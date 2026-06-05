# MS-VBAL Verification Map

Every implemented grammar rule maps to an MS-VBAL section and a test fixture.
`Status` is `Verified` (matches the cited section and has passing tests),
`Partial` (implemented with a documented, intentional deviation), or `Pending`
(not yet implemented / behavior still unconfirmed against the spec).

Spec source: see `MS-VBAL.version.md` (v20250520).

## Phase 1 - Lexer / Tokenizer

| Feature | Implementation File | Fixture | MS-VBAL Section | Status |
|---|---|---|---|---|
| Whitespace (WSC) trivia | src/analyzer/lexer/trivia.ts | tests/vbaLexer.test.ts | 3.2.2 | Verified |
| Line continuation (`_`) | src/analyzer/lexer/trivia.ts | tests/vbaLexer.test.ts | 3.2.2 | Verified |
| Line terminators (newline token) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.2.2 / 3.3.1 | Verified |
| Special / separator tokens | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.1 | Verified |
| Colon statement separator (EOS) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.1 | Verified |
| Apostrophe comments | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.1 | Verified |
| Rem comments | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.5.2 | Verified |
| Integer literals (dec/hex/octal) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.2 | Verified |
| Integer type suffixes (% & ^) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.2 | Verified |
| Float literals + suffixes (! # @) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.2 | Verified |
| Date literals (`#...#`) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.3 | Partial |
| String literals + doubled quotes | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.4 | Verified |
| Identifiers (Latin) | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.5 | Verified |
| Identifiers (non-Latin) | src/analyzer/lexer/tokenize.ts | (none yet) | 3.3.5.1 | Partial |
| Bracketed / foreign names | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.3.5.3 | Verified |
| Conditional-compilation directive marker | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.4 | Verified |

### Documented deviations

- **Date literals (Partial):** the lexer recognizes a `#...#` span on a single
  physical line as a date token but does not yet validate the inner
  `date-or-time` grammar (section 3.3.3). A `#` at statement start is treated as
  a conditional-compilation directive marker instead of a date literal.
- **Non-Latin identifiers (Partial):** Latin identifiers (section 3.3.5) are
  fully supported. Non-Latin forms (section 3.3.5.1, codepage 874/932/936/949/
  950/125x) are approximated by accepting any Unicode letter (`\p{L}`) rather
  than the exact legacy-codepage ranges. No fixtures yet.
- **Apostrophe comments:** stop at the physical line terminator (VBE behavior);
  the spec `comment-body` grammar permits embedded line-continuations, which the
  VBE does not honor.
- **Bare `&` octal:** only `&O`/`&o` (and `&H`/`&h`) begin a number; a bare `&`
  followed by digits is lexed as the concatenation operator, matching VBE usage.

## Phase 2 - Canonical Keyword Table

| Feature | Implementation File | Fixture | MS-VBAL Section | Status |
|---|---|---|---|---|
| statement-keyword | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| rem-keyword | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| marker-keyword | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| operator-identifier | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| reserved-name | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| special-form | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| reserved-type-identifier | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| literal-identifier (VBE casing) | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| future-reserved | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | 3.3.5.2 | Verified |
| reserved-for-implementation-use | src/analyzer/lexer/keywordTable.ts | (none yet) | 3.3.5.2 | Partial |
| Contextual keywords (VBE casing) | src/analyzer/lexer/keywordTable.ts | tests/vbaKeywordTable.test.ts | n/a (VBE convention) | Verified |
| Safe canonical casing edits (keyword/type/member/runtime, single-token and line-span passes) | src/analyzer/completion/canonicalCasing.ts + src/vbaMemberCompletion.ts | tests/vbaCanonicalCasing.test.ts | n/a (VBE convention) | Verified |

### Documented deviations

- **literal-identifier casing:** the spec grammar writes these lower case
  (`true` / `false` / `nothing` / `empty` / `null`); the table renders them
  capitalized (`True` / `False` / `Nothing` / `Empty` / `Null`) to match the VBE.
- **Contextual keywords:** `Explicit`, `Base`, `Compare`, `Binary`, `Text`,
  `Lib`, `Alias`, `Property`, `Step`, `Error`, `Output`, `Append`, `Random`,
  `Read`, `Object` are NOT reserved identifiers per section 3.3.5.2 but are
  capitalized by the VBE in their statement context. They are tracked separately
  and excluded from `isReservedIdentifier`.
- **reserved-for-implementation-use (Partial):** the attribute names
  (`Attribute`, `VB_Name`, ...) are listed but handled via the attribute-line
  path, not the general keyword-casing map; dotted member attribute parsing has
  focused parser fixtures.

## Phase 3 - Parser / AST

| Feature | Implementation File | Fixture | MS-VBAL Section | Status |
|---|---|---|---|---|
| Logical-statement splitting (EOS) | src/analyzer/parser/parserState.ts | tests/vbaParser.test.ts | 3.3.1 | Verified |
| Module structure | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 4.2 | Verified |
| Attribute lines, including dotted member attributes | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 4.2 | Verified |
| Option directives | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.1 | Verified |
| Module variable declarations | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3 | Verified |
| Const declarations | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.4 | Verified |
| WithEvents / New declarators | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3 | Verified |
| Declare statements with PtrSafe/Lib/Alias/params/return metadata | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3.5 | Verified |
| Conditional-compilation directives (`#Const`, `#If`, `#ElseIf`, `#Else`, `#End If`) and branch activity for symbols/diagnostics | src/analyzer/parser/parseModule.ts + src/analyzer/conditional/conditionalCompilation.ts + src/analyzer/symbols/buildModuleSymbols.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaParser.test.ts + tests/vbaConditionalCompilation.test.ts + tests/vbaSymbolGraph.test.ts + tests/vbaDiagnostics.test.ts | 3.4 | Verified |
| Type ... End Type | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3.3 | Verified |
| Enum ... End Enum | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3.4 | Verified |
| Sub / Function procedures | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.3.1 | Verified |
| Property Get / Let / Set | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.3.2 | Verified |
| Parameter lists | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.3.1.x | Verified |
| Function/Property return type | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.3.1 | Verified |
| If ... End If block | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2.1 | Verified |
| Single-line If (not a block) | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2.1 | Verified |
| Select Case ... End Select | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2.4 | Verified |
| For / For Each ... Next | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2.5 | Verified |
| Do ... Loop | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2 | Verified |
| While ... Wend | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.2 | Verified |
| With ... End With | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4.3 | Verified |
| Block-mismatch diagnostics | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.4 | Verified |
| Error recovery (never throws) | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 3.3.1 (recovery) | Verified |
| Expression AST (calls/member/ops) | src/analyzer/completion/memberAccess.ts (member-access chains only) | tests/vbaMemberCompletion.test.ts | 5.6 | Partial |
| Module symbol extraction (procs/vars/types/enums/Declares/member attributes) | src/analyzer/symbols/buildModuleSymbols.ts | tests/vbaSymbolGraph.test.ts | 5.2.3 / 5.2.3.5 / 5.2.4 / 5.3 / 4.2 | Verified |
| Project symbol graph + name resolution, including exported enum-member binding | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts | 5.3 (scope) / 5.2.3.4 (Enum) / 4.2 (visibility/default-member attributes) | Verified |
| Project visible procedure/Declare callable-name and signature surfaces | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts + tests/vbaDiagnostics.test.ts | 5.2.3.5 / 5.3 / 4.2 (visibility) | Verified |
| Project visible identifier/type-name/non-type surfaces + module/type docs + definition locations | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts + tests/vbaDiagnostics.test.ts + tests/vbaIdentifierCompletion.test.ts | 5.2.3.1.4 / 5.2.3.3 / 5.2.3.4 / 4.2; docs are n/a (XLIDE convention) | Verified |
| Project member-completion surface (object modules + visible UDT fields) + inline docs + source definition spans + `With` receiver stacks | src/analyzer/symbols/projectIndex.ts + src/analyzer/completion/memberAccess.ts + src/analyzer/hover/resolveHover.ts | tests/vbaSymbolGraph.test.ts + tests/vbaMemberCompletion.test.ts + tests/vbaDocComments.test.ts + tests/vbaHover.test.ts + tests/vbaSignatureHelp.test.ts + tests/vbaDiagnostics.test.ts | 4.2 / 5.2.3.3 / 5.3 / 5.6.9 | Partial |
| Type-name semantic tokens + hover (primitive/host/project) | src/analyzer/semantic/typeSemanticTokens.ts + src/analyzer/hover/resolveHover.ts | tests/vbaSemanticTokens.test.ts + tests/vbaHover.test.ts | 5.2.3.1.4 / 5.2.3.3 / 5.2.3.4 / 5.6.9 (`New`) | Verified |
| Duplicate-procedure detection | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts | 5.3.1 | Verified |
| Type-name completion (after `As`, creatable-only after `As New` / expression `New`) + project type docs | src/analyzer/completion/typeCompletion.ts | tests/vbaTypeCompletion.test.ts | 5.2.3.1.4 (type-spec) / 5.6.9 (`New`); docs are n/a (XLIDE convention) | Verified |
| Identifier completion (globals/code names/in-scope decls) | src/analyzer/completion/identifierCompletion.ts | tests/vbaIdentifierCompletion.test.ts | 5.6.10 (simple-name-expression) | Verified |

### Documented deviations / scope

- **Body statements (by design):** non-declaration statements inside a procedure
  body (assignments, calls, `Set`, `GoTo`, labels, `ElseIf`/`Else`/`Case`
  intermediates, etc.) are captured as a generic `Statement` node holding raw
  text. A full expression AST (section 5.6) is a later phase; the current parser
  targets reliable declaration/procedure/parameter extraction and block balance.
- **If blocks (by design):** `ElseIf` and `Else` lines are kept as ordinary body
  statements inside a single `IfBlock` rather than modeled as separate branch
  nodes. This is sufficient for block-balance diagnostics; branch modeling is
  deferred until expression parsing lands.
- **Single-line `If` detection:** an `If` opens a block only when `Then` is the
  final code token on the logical line (section 5.4.2.1); `If x Then y = 1`
  stays a single statement. Verified by fixture.
- **Conditional compilation around Declare (Partial):** external declarations and
  compiler directives are parsed, and `src/analyzer/conditional/conditionalCompilation.ts`
  can classify simple `VBA7` / `Win64` / `Win32` / `Mac` / `#Const` branches as
  active, inactive, or unknown. Analyzer branch activity defaults to modern
  Windows Office (`VBA7 = True`, `Win64 = True`, `Win32 = False`, `Mac = False`)
  and preserves explicit caller overrides. The shared branch tracker now filters
  only proven-inactive declarations/statements from the module symbol graph,
  project signatures, and active diagnostics. Remaining work is malformed
  directive-block diagnostics and broader pointer-sized API compatibility checks.
- **Recovery boundaries:** the parser recovers at newline/colon statement
  boundaries (section 3.3.1 EOS) and at module-level starters (a new
  `Sub`/`Function`/`Property`/`Type`/`Enum`/`Declare`/`Attribute`), so a missing
  `End` yields a diagnostic without swallowing the following procedure.

## Addendum - Host-Context Member Completion

The roadmap addendum requires member-completion for host-injected globals
(`ThisWorkbook.`, `Application.`, `ActiveSheet.`, `Sheet1.`, `Me.`). This needs a
host-context symbol layer + Excel object-model metadata, kept strictly separate
from the MS-VBAL core-language grammar. Tracked rows:

| Feature | Implementation File | Fixture | Source | Status |
|---|---|---|---|---|
| Host global -> type resolution table | src/analyzer/host/hostModel.ts | tests/vbaMemberCompletion.test.ts | Office VBA object model docs | Verified |
| Excel object-model member metadata | src/analyzer/host/excelObjectModel.ts + src/analyzer/host/excelReferenceMembers.ts | tests/vbaMemberCompletion.test.ts | Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30; promoted reference metadata generated from `reference/excel/json` for `Application`, `Workbook`, `Worksheet`, `Range`, `Workbooks`, `Worksheets`, and `Sheets` | Verified |
| Excel enum constants (`xlUp`, ...) | src/analyzer/host/excelReferenceMembers.ts + src/analyzer/host/hostModel.ts | tests/vbaMemberCompletion.test.ts + tests/vbaRuntime.test.ts + tests/vbaHover.test.ts + tests/vbaDiagnostics.test.ts | generated from `reference/excel/json` enum dumps, tracked in docs/excel_reference_coverage.md | Verified |
| Exhaustive `Excel.Workbook` member surface | src/analyzer/host/excelReferenceMembers.ts | tests/vbaMemberCompletion.test.ts + tests/vbaDiagnostics.test.ts + docs/excel_reference_coverage.md | `reference/excel/json/Workbook.json` plus VBE oracle `thisworkbook_unknown_member_compile` | Verified |
| Excel collection types (Workbooks/Worksheets/Sheets) + globals | src/analyzer/host/excelObjectModel.ts | tests/vbaMemberCompletion.test.ts | Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30 | Verified |
| Member-access chain resolution | src/analyzer/completion/memberAccess.ts | tests/vbaMemberCompletion.test.ts | 5.6 (member access) plus host metadata return/default-member facts, parenthesized receiver expressions, active `With` receiver stacks, and simple `Set` assignment refinement for completion | Verified |
| Worksheet code-name resolution (Sheet1) | src/analyzer/completion/memberAccess.ts + src/vbaMemberCompletion.ts | tests/vbaMemberCompletion.test.ts | Workbook project structure (listModules) | Verified |
| `Me` resolution by module kind and current source object module | src/analyzer/completion/memberAccess.ts + src/vbaMemberCompletion.ts | tests/vbaMemberCompletion.test.ts + tests/vbaHover.test.ts + tests/vbaSignatureHelp.test.ts + tests/vbaDiagnostics.test.ts | Module context | Verified |
| Typed local/param/module variable resolution | src/analyzer/completion/memberAccess.ts | tests/vbaMemberCompletion.test.ts | 5.2.3 / 5.3 (declarations) | Verified |
| Document-module event-handler completion | src/analyzer/completion/eventHandlers.ts + src/vbaMemberCompletion.ts | tests/vbaEventHandlerCompletion.test.ts + tests/vbaDiagnostics.test.ts | Excel event signatures + workbook/document module context | Partial: workbook and worksheet handlers plus wrong-module guidance verified; chart/UserForm designer-backed handlers pending |
| VS Code completion provider (trigger `.`) | src/vbaMemberCompletion.ts | (manual) | n/a | Verified |

Verification rule for this addendum: VBA language grammar stays verified against
MS-VBAL; host object-model members must be verified against the Office VBA
object-model documentation, generated COM type-library metadata, or recorded VBE
behavior fixtures. LLM-generated member lists are never authoritative.

## Addendum - Built-In VBA Runtime Metadata

Phase 9 adds hover and identifier-completion for the intrinsic VBA runtime
library (`MsgBox`, `Left`, `CLng`, `Now`, `Array`, `RGB`, ...). These are core
language built-ins (the VBA standard library), distinct from host object-model
members. The metadata is a typed TS module rather than the JSON file the roadmap
originally suggested, matching the host-model precedent for compile-time checking.

| Feature | Implementation File | Fixture | Source | Status |
|---|---|---|---|---|
| Built-in runtime function/statement signatures | src/analyzer/runtime/vbaRuntime.ts | tests/vbaRuntime.test.ts | learn.microsoft.com/office/vba/language + MS-VBAL | Verified |
| Built-in runtime constants (`vbOKOnly`, `vbCrLf`, ...) | src/analyzer/runtime/vbaRuntime.ts | tests/vbaRuntime.test.ts + tests/vbaHover.test.ts + tests/vbaDiagnostics.test.ts | learn.microsoft.com/office/vba/language + `reference/vba/json` | Verified |
| Runtime hover resolution (after user/host/code-name) | src/analyzer/hover/resolveHover.ts | tests/vbaHover.test.ts | n/a | Verified |
| Runtime identifier completion (`runtime` function kind plus constant completions) | src/analyzer/completion/identifierCompletion.ts | tests/vbaRuntime.test.ts | n/a | Verified |

Verification rule: built-in signatures and constants must be transcribed from
the Microsoft VBA language reference, MS-VBAL, or checked-in reference dumps,
never invented. Names that collide with intrinsic data types (`Date`, `Time`,
`String`, `Error`) are deliberately omitted to avoid type/function ambiguity in
`As` positions.

---

## Addendum - Signature Help (Parameter Info)

Phase 7/9 adds the VBE call tip: when the caret is inside a call's argument
list, `src/analyzer/signature/signatureHelp.ts` returns the callee's signature
and the active-parameter index. The signature comes from one of three verified
sources (host members, user procedures/Declares, runtime built-ins); no
signature is ever invented, so an unknown callee yields no tip.

| Feature | Implementation File | Fixture | Source | Status |
|---|---|---|---|---|
| Caret-to-active-call resolution (paren + parenless) | src/analyzer/signature/signatureHelp.ts | tests/vbaSignatureHelp.test.ts | MS-VBAL 5.4.2 (call statements) | Verified |
| Host-member call signatures (`Workbooks.Open`, `Range.Offset`, ...) | src/analyzer/host/excelObjectModel.ts (`memberSignatures`) | tests/vbaSignatureHelp.test.ts | Office VBA object-model reference (learn.microsoft.com) | Verified |
| User procedure/Declare signatures and Declare Lib/Alias call-tip details (from AST) | src/analyzer/signature/signatureHelp.ts | tests/vbaSignatureHelp.test.ts | n/a (built from parsed `ProcedureNode` / `DeclareNode`; Declare grammar row above covers MS-VBAL 5.2.3.5) | Verified |
| Runtime built-in signatures | src/analyzer/runtime/vbaRuntime.ts | tests/vbaSignatureHelp.test.ts | learn.microsoft.com/office/vba/language + MS-VBAL | Verified |

Verification rule: host-member signatures are transcribed from the Office VBA
object-model reference. Where a method has a large variadic tail (e.g.
`Application.Run` takes Arg1..Arg30) only the leading commonly-used parameters
are listed rather than inventing a synthetic `...` token.

---

## Addendum - VBA7 / PtrSafe Completion Snippets

Phase 6 adds migration snippets for common external declaration patterns. These
snippets are editor affordances rather than diagnostics: they follow Microsoft
Learn's `PtrSafe` guidance that VBA7 compatibility uses `#If VBA7 Then ...
#Else ... #End If`, that 64-bit Office `Declare` statements require `PtrSafe`,
and that pointer/handle-sized parameters or returns should use `LongPtr`.

| Feature | Implementation File | Fixture | Source | Status |
|---|---|---|---|---|
| Access-modifier `Declare PtrSafe` Sub/Function snippets | src/analyzer/completion/keywordCompletion.ts | tests/vbaKeywordCompletion.test.ts | Microsoft Learn `PtrSafe keyword (VBA)` | Verified |
| `#If VBA7` / `#If Win64` conditional-compilation snippets with `Declare PtrSafe` templates | src/analyzer/completion/keywordCompletion.ts | tests/vbaKeywordCompletion.test.ts | Microsoft Learn `PtrSafe keyword (VBA)` + `Compiler constants (VBA)` | Verified |

---

## Addendum - Active Diagnostics

Phase 5 adds live semantic diagnostics computed from module text by
`src/analyzer/diagnostics/analyzeModule.ts` (rule catalogue in
`ruleMetadata.ts`). Each rule is high-confidence and cites the MS-VBAL section
it enforces; the engine is merged with the structural block-balance analyzer
(`src/vbaStructuralAnalysis.ts`, which covers the "Missing End .../unexpected terminator"
family) in `registerVbaDiagnostics`.

| Rule code | Meaning | MS-VBAL | Implementation File | Fixture | Status |
|---|---|---|---|---|---|
| `unterminated-string` | String literal with no closing quote | 3.3.4 (string literal token) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-line-continuation` | A likely line-continuation underscore is missing required leading whitespace, or has trailing text/comment before the physical line ends | 3.2.2 (line-continuation) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-procedure` | Two procedures share a name (Get/Let/Set excepted) | 5.3 (procedure declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-declaration` | Param/local redeclared in one procedure scope | 5.2 / 5.3 (declared names) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-module-variable` | Module-level variable redeclared | 5.2.3 (module variable declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-enum-member` | The same case-insensitive member name appears more than once inside one active top-level `Enum` block; inactive whole-Enum conditional branches stay quiet, while conditional directives nested inside an `Enum` remain deferred until parser support is explicit | 5.2.3.4 (Enum member declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `ambiguous-enum-member` | A value-position bare read resolves to multiple visible Enum member definitions across distinct `Enum` containers; same-module bindings and procedure locals/parameters shadow exported ambiguity, and declaration-only same-name members across separate Enums stay quiet per the accepted oracle control | VBE oracle: Ambiguous name detected | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `const-assignment` | Assignment to a declared `Const` | 5.4.3.1 (Const cannot be assigned) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `declare-missing-ptrsafe` | An active `Declare` statement lacks `PtrSafe` while supplied compiler constants prove `Win64`; unknown platform environments and inactive legacy branches remain silent | VBA 7 64-bit Office `PtrSafe` requirement | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `option-explicit-missing` | Code module omits `Option Explicit` (configurable) | 5.2.4.1.1 (Option Explicit) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `undeclared-variable` | `Option Explicit` module writes to or reads an identifier in a high-confidence value position that resolves to no local/module/project/runtime/host identifier; covered positions include bare assignment/`Set` targets, RHS and call-argument reads, module-qualified project procedure/value qualifiers, known standard-module qualifiers even when the member is absent, control-flow block headers, member receivers, and indexed bases, while type-name, label, named-argument, and unresolved external-call positions are skipped; runtime constants (`vbOKOnly`) and generated Excel enum constants (`xlUp`) suppress false positives; missing `Option Explicit` remains implicit Variant | 5.2.4.1.1 (Option Explicit) | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/symbols/projectIndex.ts | tests/vbaDiagnostics.test.ts + tests/vbaSymbolGraph.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `unknown-call` | Call statement whose callee is a bare (non-member) identifier - lone identifier, parenless args (`MsgBox "hi"`), or `Call Foo` - that resolves to no project procedure/Declare, runtime function, host global, `Application` member, or in-scope name | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-proc-header` | A `Sub`/`Function`/`Property` header where a token other than `(` (or `As` for a `Function`/`Property Get`) follows the procedure name (e.g. `Sub My Sub`) | 5.3.1 (procedure declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-identifier-start` | Declaration names that begin with a digit, such as `Dim 1bad As Long`; active coverage is limited to declaration contexts and leaves leading-underscore, Unicode, and illegal-character identifier edge cases deferred | 3.3.5 (lex-identifier) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `module-declaration-in-procedure` | Module-only forms inside procedure bodies: semantic coverage for completed statement forms such as `Option`, `Attribute` lines outside exported-source metadata mode, `Def*`, and `Public`/`Private`/`Friend`/`Global` declarations; structural coverage for indented nested `Type`, `Enum`, `Declare`, and procedure declarations that would otherwise look block-balanced or break parser recovery. VBE `CodeModule.AddFromString`/live-style probes reject visible member `Attribute` lines as Syntax error; unindented exported-source member `Attribute` metadata targeting the current procedure remains accepted by XLIDE's exported-source parser policy only after a module-level `VB_` `Attribute` marker such as `Attribute VB_Name` is present | 5.2 (module/procedure metadata and declarations) / 5.3 (procedure declarations) / VBE oracle | src/analyzer/parser/parseModule.ts + src/analyzer/diagnostics/analyzeModule.ts + src/vbaStructuralAnalysis.ts | tests/vbaDiagnostics.test.ts + tests/vbaStructuralAnalysis.test.ts + tests/vbaModuleAnalysis.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `module-declaration-after-procedure` | A parsed top-level module declaration appears after an active procedure declaration: `Declare`, `Event`, module-level `Const`/variable groups, `Type`, `Enum`, and `Def*` declarations must be in the module declarations section before procedures. Later procedures remain accepted; inactive conditional branches stay silent; a valid `#If`/`#ElseIf`/`#Else` block after a procedure reports only the active declaration branch; declarations inside a malformed conditional-compilation branch-order block are suppressed as cascade noise because `else-branch-order` owns the primary compile error | 5.2 (module body structure) / 5.3 (procedure declarations) / VBE oracle | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `statement-outside-procedure` | An executable, call, assignment, or expression statement appears as an active top-level module member instead of inside a `Sub`, `Function`, or `Property` procedure; declaration-section statement forms such as `Def*` stay accepted before procedures, and `Implements` remains owned by `implements-statement-placement` | 5.2 (module body structure) / 5.4 (statements) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-declaration-name` | An unbracketed reserved identifier such as `Dim` or `In` is used where a procedure, variable, parameter, Type, Enum, field, or enum-member declaration name must be an IDENTIFIER; bracketed FOREIGN-NAME forms are accepted | 3.3.5.2 (reserved identifiers) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `unbalanced-parens` | A `(` left open at a statement boundary, or a `)` with no matching `(`, within one logical statement | 3.3.1 (special tokens) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `argument-count` | A call statement to a same-module, unique exported project, module-qualified exported standard-module Sub/Function/Declare, or verified runtime function with an explicit parameter-list signature supplies too few/too many arguments (Optional/ParamArray aware), a named argument names no parameter, or a valid source-backed/host member-call context, parenthesized or parenless, violates a known member signature. Bare runtime signatures are skipped when a visible non-callable source name shadows the runtime name | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `argument-type-mismatch` | A same-module, unique exported project, module-qualified exported standard-module, Declare, curated runtime, or known source-backed/host member call receives an argument whose inferred type is a provable deterministic runtime type error; argument inference includes literals, variables, nested calls, zero-argument Function/Property Get value references, module-qualified project Function references, numeric/string expressions, and known member returns. Focused oracle cases compile successfully but runtime probes raise error 13 for nonnumeric string-to-numeric coercion while numeric-string controls run, and raise error 6 for decimal literals outside Byte/Integer bounds while min/max boundary controls run. Bare runtime parameter/return metadata is skipped when visible source names shadow the runtime name | 5.3.1 / 5.2.3.5 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `argument-object-type-mismatch` | A same-module, curated runtime, source-backed member, or host/reference member call receives a scalar argument where an object parameter is required; argument inference includes zero-argument Function/Property Get value references, module-qualified project Function references, and proven source-backed/host member return expressions. Runtime return metadata remains runtime-last and is skipped for source-shadowed bare names. Error severity because a focused VBE oracle case rejects scalar-to-object arguments at compile time | 5.3.1 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `byref-argument-type-mismatch` | A source-backed same-module or unique exported project call passes an unparenthesized known scalar variable whose type differs from a ByRef parameter type; focused VBE oracle cases show `ByRef Long` accepts a `Long` variable, a literal, and a parenthesized `Long` variable expression, rejects an unparenthesized `Integer` variable as `ByRef argument type mismatch`, and the matching `ByVal Long` control accepts an `Integer` variable | 5.3.1 / VBE oracle: ByRef argument type mismatch | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `assignment-type-mismatch` | A scalar assignment, including a Function/Property Get return name or source-backed writable workbook class property/public field, receives a value whose inferred type is a provable deterministic runtime type error; focused oracle cases compile successfully but runtime probes raise error 13 for nonnumeric string-to-numeric/Boolean/property/public-field coercion while valid coercion controls run, and raise error 6 for decimal literals outside Byte/Integer bounds while min/max boundary controls run | 5.4.3 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `missing-return-assignment` | A closed untyped `Function` or `Property Get` has no assignment to its return variable anywhere in the active body; typed returns and parser-recovered conditional header forms stay silent because explicit assignment is not required for VBA to return the declared type's default value and XLIDE must not guess at recovered compiled bodies. Warning severity because VBA falls through to the default value rather than raising a compile error | VBA Function return variable semantics / XLIDE type-safety policy | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `assignment-object-type-mismatch` | A `Set` assignment targets a known object variable, Function/Property Get return name, or source-backed object-valued member, and the RHS is a deterministic incompatible object/scalar value, including scalar zero-argument Function/Property Get value references and proven source-backed/host member return expressions. Host collection property calls with arguments, such as `Worksheet.ListObjects("Tests")`, use default `Item` return inference only when metadata proves the returned collection exposes `Item` and the property itself has no parameter signature; same project class and explicit `Implements` assignments are accepted | 5.4.3 / Set statement | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/symbols/projectIndex.ts | tests/vbaDiagnostics.test.ts + tests/vbaSymbolGraph.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `readonly-member-assignment` | A source-backed workbook class property assignment targets a member whose source surface has no setter; focused VBE oracle evidence rejects this as `Can't assign to read-only property` | 5.4.3 / VBE oracle read-only property | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `member-not-found` | A source-backed workbook class/UDT/standard-module receiver or dump-backed exhaustive host receiver uses a member name absent from the known member surface; known standard-module qualifiers are source-exhaustive, including empty visible surfaces where all candidates are private. Focused VBE oracle evidence rejects unknown class property/method cases, `ThisWorkbook.DoesntExist`, and calling workbook events such as `ThisWorkbook.AfterSave` as `Method or data member not found`, while known property/public-field controls compile and generated exhaustive `Excel.Workbook`/`Excel.Worksheet`/`Excel.Range` coverage handles host cases such as `ActiveSheet.asdf`, declared `Worksheet` receivers, chained `Workbooks(1).Worksheets(1)` receivers, `ActiveCell`, declared `Range`, chained `Worksheet.Range(...)` receivers, parenthesized receivers such as `(p.Child).Missing`, and leading-dot receivers inside nested `With .Member` stacks. Late-bound `Object`/`Variant` receivers suppress hard absence diagnostics even after simple `Set` assignments to known host objects because VBE compile controls accept unknown members there | 5.6.9 / VBE oracle member binding plus generated Excel reference metadata | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects + syntax_corpus/oracle/vbe_oracle_cases.json + docs/excel_reference_coverage.md | Verified |
| `string-arithmetic-coercion` | A numeric context contains an arithmetic expression with a provably nonnumeric string literal; error severity because focused VBE oracle cases show it compiles but deterministically raises runtime error 13 | 5.6 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `division-by-zero` | A settled procedure statement contains `/`, integer division `\`, or `Mod` with a decimal/hex/octal numeric literal zero divisor, a zero-valued same-module/procedure integer `Const`, zero-valued current-module Enum member, bare or module-qualified visible exported standard-module `Const`/Enum member when workbook project context is available, or a parenthesized integer expression over those constants that folds to zero. Exported standard-module constants may fold through private same-module integer helper constants without making those helpers visible to callers. Error severity because focused VBE oracle cases show these expressions compile but deterministically raise runtime error 11, while nonzero literal, Const, and Enum member controls run | 5.6 / runtime division by zero | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/symbols/projectIndex.ts | tests/vbaDiagnostics.test.ts + tests/vbaSymbolGraph.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `runtime-argument-value` | A settled procedure statement calls `Left`, `Left$`, `Right`, `Right$`, `String`, `String$`, `Space`, `Space$`, `Mid`, `Mid$`, `Replace`, three-or-more-argument `InStr`, `Chr`, or `ChrW` with an integer value outside the known runtime bound: non-negative Length/Number arguments, `Mid`/`Mid$` Start less than 1, `Replace` Start less than 1, `Replace` Count less than -1, `InStr` Start less than 1, `Chr` CharCode outside `0..255`, or `ChrW` CharCode greater than `65535`. Values may be integer literals, reducible integer expressions, same-module/procedure `Const` values, current-module Enum members, or bare/module-qualified visible exported standard-module `Const`/Enum members when workbook project context is available; exported constants may fold through private same-module integer helper constants. Error severity because focused VBE oracle cases show these calls compile but deterministically raise runtime error 5, while in-bound controls run. Bare calls are skipped when source symbols or local declarations shadow the intrinsic; explicit `VBA.Left$`/`VBA.Chr` remain intrinsic-bound | 5.6 / runtime argument bounds | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/symbols/projectIndex.ts | tests/vbaDiagnostics.test.ts + tests/vbaSymbolGraph.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `dim-initializer` | A variable declaration includes a VB.NET-style inline initializer (`Dim x As Long = 1`), which VBA does not allow; `Const` is exempt | 5.2.3.1 (variable declaration) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `fixed-array-redim` | A `ReDim` or `ReDim Preserve` target resolves in the active procedure/module declaration scope to a fixed-size array declaration such as `Dim Values(1 To 3) As Long`; dynamic arrays, undeclared `ReDim` targets, local dynamic shadows of module fixed arrays, and inactive fixed declarations stay silent | ReDim statement / fixed-size array declaration semantics | src/analyzer/parser/parseModule.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaParser.test.ts + tests/vbaDiagnostics.test.ts | Verified |
| `redim-preserve-dimension-change` | A straight-line active `ReDim Preserve` target has a prior known `ReDim` shape and changes a comparable non-final dimension, the known dimension count, or a comparable final-dimension lower bound; active coverage compares literal-style bounds and does not propagate shapes learned inside nested blocks outward. Focused runtime oracle fixtures show resizing a non-final dimension and changing the final dimension lower bound raise runtime errors while resizing only the final dimension upper bound runs successfully | ReDim Preserve deterministic runtime behavior | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `invalid-erase-target` | An `Erase` target-list entry is clearly not a variable/array target name, such as a literal or arithmetic expression (`Erase 1 + 2`); variable-like unresolved targets remain quiet until array-ness/type binding is modeled more deeply | Erase statement target syntax | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `type-declaration-character-as-clause` | A declaration name uses a legacy type-declaration character (`$`, `%`, `&`, `!`, `#`, `@`, or `^`) and also has an explicit `As` clause in a VBE-rejected declaration surface: variable declarations, Const declarations, parameters, UDT fields, and Functions. Suffix-only declarations normalize to the base name and inferred suffix type for variables/constants, parameters, UDT fields, Functions, Property Gets, and Declare Functions; VBE-verified controls show `Property Get Name$() As String` declarations compile | 5.2.3.1 / 5.3.1 type-declaration characters / VBE oracle | src/analyzer/parser/parseModule.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaParser.test.ts + tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `unexpected-declaration-token` | A declaration, parameter, or UDT field has an extra same-statement token after a complete `As` type name, such as `Dim s As String junk`; recognized fixed-length string suffixes (`As String * n`) are consumed before trailing-token detection, so `Dim s As String * 10 junk` reports the extra token after the suffix. Focused VBE oracle evidence rejects the representative `Dim` case as `Syntax error` | 5.2.3.1 / VBE oracle Syntax error | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `fixed-length-string-size` | A fixed-length `String` declaration size is outside the VBE-verified `1..65526` range when the suffix is a decimal integer literal or a same-module/procedure `Const`/current-module Enum member reducible to a deterministic integer expression (`+`, `-`, `*`, unary signs, parentheses, decimal/hex/octal integer literals, and resolvable `Const`/Enum references). `String * 1` and `String * 65526` compile, while `String * 0` and `String * 65527` are rejected as `Invalid length for fixed-length string`. Unknown, duplicate, and non-deterministic expressions remain deferred | MS-VBAL fixed-length String bounds / VBE oracle | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `object-module-public-member` | A class/document/UserForm module declares explicit Public constants, arrays, fixed-length strings, user-defined Types, or Declare statements; focused VBE oracle evidence rejects each branch as an invalid public object-module member while standard modules remain outside this rule | VBE oracle object-module public member restrictions | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `event-declaration-module-kind` | A standard module declares `Event`; `Event` declarations are object-module declarations, while class/document/UserForm module declarations remain accepted and inactive conditional branches are skipped | 5.2.5 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects/workbook_analysis_module_kind_canary.json | Verified |
| `withevents-declaration` | A `WithEvents` variable is declared in a standard module, inside a procedure body, `As New`, or as an array; module-level class/document/UserForm `WithEvents` variables remain accepted, inactive conditional branches are skipped, and event-source type compatibility such as generic `Object` remains deferred | 5.2.3 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects/workbook_analysis_module_kind_canary.json | Verified |
| `friend-declaration` | A `Friend` procedure is declared in a standard module, or `Friend` is used as a module-variable modifier; object-module `Friend` procedures remain accepted, inactive conditional branches are skipped, and broader `Friend` visibility/binding semantics remain deferred | 5.3 procedure visibility / Friend keyword | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects/workbook_analysis_module_kind_canary.json | Verified |
| `implements-statement-placement` | An `Implements` statement is declared in a standard module, inside a procedure body, or after a procedure in an object module; object-module declaration-section statements remain accepted, inactive conditional branches are skipped, and interface member completeness remains deferred to the project binder | Implements statement | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects/workbook_analysis_module_kind_canary.json | Verified |
| `raiseevent-undeclared-event` | A settled `RaiseEvent` statement names no active `Event` declaration in the containing module; declared same-module events, inactive `RaiseEvent` statements, and inactive `Event` declarations are handled through conditional-compilation activity. Event argument count/type validation remains deferred | RaiseEvent statement | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + tests/fixtures/vbaProjects/workbook_analysis_module_kind_canary.json | Verified |
| `event-handler-module-scope` | A `Sub` name matches a known Excel workbook/worksheet event handler but is declared outside the matching document-module context; information severity because the procedure may still compile as an ordinary procedure while not being wired as an event handler | Excel document-module event binding | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/completion/eventHandlers.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-as-type-name` | A shared type-name position (`As`, return, parameter, UDT field, `New`, `TypeOf ... Is`, or `Implements`) uses an unresolved reserved identifier/runtime function, a visible project declaration known not to be a type, or an ambiguous visible project type name. Unknown external/reference-library names remain deferred | 3.3.5.2 / 5.2.3.1 / 5.6.9 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `invalid-new-type-name` | A resolved non-creatable type name, such as a primitive, Excel host object type, document module, Enum, or UDT, is used after expression-level `New` or declaration-level `As New`; unresolved external/reference-library names remain deferred | 5.2.3.1 / 5.6.9 (`New`) | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/completion/typeCompletion.ts | tests/vbaDiagnostics.test.ts + tests/vbaTypeCompletion.test.ts | Verified |
| `set-required` | A plain assignment targets a known object variable, Function/Property Get return name, or source-backed object-valued member (`Property Set`/public field), where VBA requires `Set` | 5.4.3 / Set statement | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `set-requires-object` | A `Set` assignment targets a known intrinsic scalar variable or source-backed scalar member; focused VBE oracle evidence rejects scalar `Integer` and scalar `String` targets, including `Set text = New Collection` where the right-hand side is object-valued | 5.4.3 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `scalar-member-access` | A member-access dot targets a variable declared as a known intrinsic scalar (`String`, numeric, `Boolean`, `Date`); named scalar members are VBE Compile `Invalid qualifier` errors and trailing scalar dots are VBE Compile `Syntax error`s | 5.6.9 / VBE oracle Invalid qualifier and Syntax error | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `call-requires-parens` | A `Call` statement supplies arguments without enclosing parentheses (`Call MsgBox "hi"`) | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-explicit-call-target` | A runtime entry with verified explicit-`Call` incompatibility is used as the target of `Call`; VBE rejects both `Call DoEvents` and `Call DoEvents()` as Syntax error while accepting bare `DoEvents` and expression `DoEvents()` controls. Source/project callables and visible source shadows suppress the runtime-only special case | VBE oracle call syntax | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/runtime/vbaRuntime.ts | tests/vbaDiagnostics.test.ts + tests/vbaRuntime.test.ts + tests/vbaSignatureHelp.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `call-statement-forbids-parens` | A standalone zero-argument call statement uses empty parentheses without `Call`: known same-module/exported project procedures/Declares such as `myFunction()`, verified zero-argument runtime functions such as `DoEvents()`, plus member/property statements such as `ThisWorkbook.CanCheckIn()`, `Application.Calculate()`, and `ActiveSheet.Range()`. Required-argument calls such as `MsgBox()` remain owned by `argument-count`, and source-shadowed runtime names stay silent on runtime-only syntax. Expression context and non-empty standalone member/property controls are VBE-oracle verified separately; explicit `Call` target special cases are owned by `invalid-explicit-call-target` | 5.4.2.1 (call statement) / VBE oracle call syntax | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `invalid-expression-syntax` | A narrow expression-syntax diagnostic for impossible operator shapes: consecutive non-unary binary operators such as `***`, a statement ending in a binary operator, unsupported `?` conditional-operator syntax, and oracle-backed incomplete member access. `IIf(...)` remains the supported inline conditional function form. Broader incomplete-expression analysis is deferred to avoid noisy realtime false positives | 5.6 / VBE oracle Syntax error behavior | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `required-param-after-optional` | A required parameter follows an `Optional` parameter in a procedure header | 5.3.1.5 (parameter list) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `parameter-array-as-type-syntax` | A procedure parameter places array parentheses after the `As` type, such as `ByVal values As Long()`. VBA parameter arrays use `values() As Long`; this specific diagnostic supersedes the generic trailing-token recovery diagnostic for the misplaced `()` | 5.3.1.5 (parameter list) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `paramarray-not-last` | A `ParamArray` parameter is not the final parameter | 5.3.1.6 (ParamArray) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `paramarray-with-optional` | A procedure parameter list combines `ParamArray` with one or more `Optional` parameters; `ParamArray` must stand alone as the variadic tail and cannot be used in the same parameter list as Optional arguments | 5.3.1.5 / 5.3.1.6 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `paramarray-non-variant` | A `ParamArray` parameter explicitly declares a non-Variant element type such as `ParamArray values() As String`; `ParamArray` elements must be Variant. Omitted `As` type and explicit `As Variant` remain accepted | 5.3.1.6 (ParamArray) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `property-setter-missing-value` | A `Property Let` or `Property Set` declaration has no parameter list value slot, such as `Property Let Name()`. Setters must include a final value parameter; `Property Get` remains outside this rule | 5.3.1.4 (Property procedures) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `property-set-scalar-value` | A `Property Set` declaration's final value parameter is declared as a known intrinsic scalar such as `Long` or `String`; `Property Set` assigns object references, so unknown/project/host object-looking types stay outside this scalar-only rule | 5.3.1.4 (Property Set procedures) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `property-let-object-value` | A `Property Let` declaration's final value parameter resolves to a known object-reference type such as `Object`, an Excel host object, or a project class/document/UserForm; object assignment belongs to `Property Set`, while unresolved or ambiguous type names stay deferred | 5.3.1.4 (Property Let procedures) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `property-setter-return-type` | A `Property Let` or `Property Set` declaration includes a return `As` clause after its parameter list, such as `Property Let Name(ByVal value As Long) As Long`; only `Property Get` declares a property return type | 5.3.1.4 (Property procedures) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `property-accessor-signature-mismatch` | A paired `Property Let` or `Property Set` declaration's leading index parameters are incompatible with the same-name `Property Get` parameters before the setter's final value parameter; active coverage checks count, array shape, effective passing mode, and known scalar/Variant type mismatches while leaving unresolved object type aliasing for later binder work | 5.3.1.4 (Property procedures) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `parameter-default-type-mismatch` | An Optional parameter default is a deterministic string literal that VBE rejects for the declared scalar type; focused oracle cases show numeric and numeric-string defaults for `Long`, string defaults for `String`, and `True` for `Boolean` compile, while nonnumeric string defaults for `Long` and `Boolean` reject as `Type mismatch` | 5.3.1 / VBE oracle: Type mismatch | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `exit-wrong-proc` | An `Exit Sub`/`Exit Function`/`Exit Property` does not match the enclosing procedure kind (`Exit Do`/`Exit For` excluded) | 5.4.1.3 (Exit statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `else-branch-order` | `#Else` and normal `Else` are final peer branches: `#ElseIf`/`ElseIf` after `#Else`/`Else` and duplicate `#Else`/`Else` branches now produce red diagnostics. Conditional-compilation directive ordering is checked structurally regardless of branch activity; normal `If` blocks respect inactive conditional-compilation regions; nested blocks keep independent branch state | 3.4 (conditional compilation) / 5.4.2.1 (If block) / VBE oracle | src/analyzer/conditional/conditionalCompilation.ts + src/analyzer/parser/parseModule.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `duplicate-label` | A named label or normalized decimal line label is declared more than once in the same active procedure body; label scopes remain procedure-local and inactive conditional-compilation branches are ignored | 5.4.1 (procedure labels) | src/analyzer/diagnostics/analyzeModule.ts + src/analyzer/flow/procedureLabels.ts | tests/vbaDiagnostics.test.ts | Verified |
| `next-variable-mismatch` | A `Next name` closer supplies a simple control variable that does not match the active `For` or `For Each` opener's simple control variable; omitted `Next` variables, matching names, nested loops, inactive closer branches, and complex/recovered variable shapes stay silent | 5.4.2.5 (For / For Each ... Next) | src/analyzer/parser/parseModule.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaParser.test.ts + tests/vbaDiagnostics.test.ts | Verified |
| `for-each-control-variable-type` | A `For Each` control variable must be `Variant` or an object variable; parser-backed loop metadata plus declaration-shape binding now reject known intrinsic scalars, array variables, and project UDT/Enum controls when type metadata is available. Variant, Object, host/project object-looking types, implicit Variant, ambiguous types, and inactive branches stay quiet | 5.4.2.5 (For Each ... Next) / VBE oracle | src/analyzer/parser/parseModule.ts + src/analyzer/symbols/buildModuleSymbols.ts + src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `option-after-declaration` | An `Option` statement appears after a declaration or procedure (only `Attribute` lines may precede it) | 5.2.1 (module options) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |

Deliberately deferred (not shipped): full flow-aware identifier binding,
ambiguous external-reference behavior, and the broad arbitrary-expression form
of `unknown-call`. The `undeclared-variable` rule now ships for project-backed
`Option Explicit` write/read positions, including bare assignment/`Set` targets,
RHS and call-argument reads, control-flow block headers, member receivers, and
indexed bases, while skipping type-name, label, named-argument, and unresolved
external-style call positions. The `unknown-call` rule now ships for the three
unambiguous call forms - a lone identifier, a parenless call with arguments
(`msrbox ""`), and an explicit `Call` - while the implicit-host-member form
`Cells(1, 1)` / `Range("A1")` and any statement containing a top-level `=`
(assignment) are excluded. Empty-parentheses standalone calls without `Call`
are handled first by `call-statement-forbids-parens` when the target is a known
zero-argument same-module/exported project procedure, verified runtime function,
or known member-call shape.
Argument-count validation (`argument-count`) is likewise limited to
same-module, deterministic project signature calls, verified runtime signatures
with explicit parameter lists, and valid parenthesized or parenless member-call
contexts whose source/host metadata provides a known signature: ambiguous bare exported
project names stay silent, while module-qualified standard-module calls resolve
through the named module only.
Argument and assignment type diagnostics ship only where the local expression
model can infer both sides deterministically and the behavior is backed by
compile/runtime oracle evidence; broad object/member type checking beyond known
signatures stays deferred. The
remaining deferred cases require a full
expression binder plus a complete host catalogue; without them they would emit
false positives, which the project's no-false-positive rule forbids. They will
ship only once they can be proven safe. The shipped `invalid-line-continuation`
rule is intentionally limited to settled malformed physical-line shapes; dangling
final continuations remain a separate realtime-recovery/save-validation policy.

Verification rule: a new diagnostic rule must (1) carry an MS-VBAL
`specReference` in `ruleMetadata.ts`, (2) be high-confidence, and (3) have
positive and negative fixtures in `tests/vbaDiagnostics.test.ts`.
