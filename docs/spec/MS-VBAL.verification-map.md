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
| Conditional-compilation directives | src/analyzer/lexer/tokenize.ts | tests/vbaLexer.test.ts | 3.4 | Partial |

### Documented deviations

- **Date literals (Partial):** the lexer recognizes a `#...#` span on a single
  physical line as a date token but does not yet validate the inner
  `date-or-time` grammar (section 3.3.3). A `#` at statement start is treated as
  a conditional-compilation directive marker instead of a date literal.
- **Non-Latin identifiers (Partial):** Latin identifiers (section 3.3.5) are
  fully supported. Non-Latin forms (section 3.3.5.1, codepage 874/932/936/949/
  950/125x) are approximated by accepting any Unicode letter (`\p{L}`) rather
  than the exact legacy-codepage ranges. No fixtures yet.
- **Conditional-compilation directives (Partial):** the `#` marker is tokenized
  and the following `If` / `Else` / `ElseIf` / `Const` / `End` lex as keywords,
  but directive blocks are not yet parsed (deferred to a later phase).
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
  path, not the general keyword-casing map; no fixtures yet.

## Phase 3 - Parser / AST

| Feature | Implementation File | Fixture | MS-VBAL Section | Status |
|---|---|---|---|---|
| Logical-statement splitting (EOS) | src/analyzer/parser/parserState.ts | tests/vbaParser.test.ts | 3.3.1 | Verified |
| Module structure | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 4.2 | Verified |
| Attribute lines | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 4.2 | Verified |
| Option directives | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.1 | Verified |
| Module variable declarations | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3 | Verified |
| Const declarations | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.4 | Verified |
| WithEvents / New declarators | src/analyzer/parser/parseModule.ts | tests/vbaParser.test.ts | 5.2.3 | Verified |
| Declare statements | src/analyzer/parser/parseModule.ts | (none yet) | 5.2.3.5 | Partial |
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
| Module symbol extraction (procs/vars/types/enums) | src/analyzer/symbols/buildModuleSymbols.ts | tests/vbaSymbolGraph.test.ts | 5.2.3 / 5.2.4 / 5.3 | Verified |
| Project symbol graph + name resolution | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts | 5.3 (scope) / 4.2 (visibility) | Verified |
| Project visible procedure-name surface | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts + tests/vbaDiagnostics.test.ts | 5.3 / 4.2 (visibility) | Verified |
| Project visible type-name surface | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts | 5.2.3.1.4 / 5.2.3.3 / 5.2.3.4 / 4.2 | Verified |
| Project class member-completion surface + inline docs | src/analyzer/symbols/projectIndex.ts + src/analyzer/completion/memberAccess.ts + src/analyzer/hover/resolveHover.ts | tests/vbaSymbolGraph.test.ts + tests/vbaMemberCompletion.test.ts + tests/vbaDocComments.test.ts | 4.2 / 5.3 / 5.6.9 | Partial |
| Project type semantic tokens | src/analyzer/semantic/typeSemanticTokens.ts | tests/vbaSemanticTokens.test.ts | 5.2.3.1.4 / 5.2.3.3 / 5.2.3.4 / 5.6.9 (`New`) | Verified |
| Duplicate-procedure detection | src/analyzer/symbols/projectIndex.ts | tests/vbaSymbolGraph.test.ts | 5.3.1 | Verified |
| Type-position completion (after As / As New) | src/analyzer/completion/typeCompletion.ts | tests/vbaTypeCompletion.test.ts | 5.2.3.1.4 (type-spec) | Verified |
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
- **Declare (Partial):** the procedure name, Sub/Function kind, and visibility
  are extracted; the `Lib`/`Alias`/parameter detail of external declarations
  (section 5.2.3.5) is not yet modeled. No dedicated fixture.
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
| Excel object-model member metadata | src/analyzer/host/excelObjectModel.ts | tests/vbaMemberCompletion.test.ts | Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30; promoted reference metadata generated from `reference/excel/json` | Verified |
| Exhaustive `Excel.Workbook` member surface | src/analyzer/host/excelReferenceMembers.ts | tests/vbaMemberCompletion.test.ts + tests/vbaDiagnostics.test.ts + docs/excel_reference_coverage.md | `reference/excel/json/Workbook.json` plus VBE oracle `thisworkbook_unknown_member_compile` | Verified |
| Excel collection types (Workbooks/Worksheets/Sheets) + globals | src/analyzer/host/excelObjectModel.ts | tests/vbaMemberCompletion.test.ts | Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30 | Verified |
| Member-access chain resolution | src/analyzer/completion/memberAccess.ts | tests/vbaMemberCompletion.test.ts | 5.6 (member access) | Verified |
| Worksheet code-name resolution (Sheet1) | src/analyzer/completion/memberAccess.ts + src/vbaMemberCompletion.ts | tests/vbaMemberCompletion.test.ts | Workbook project structure (listModules) | Verified |
| `Me` resolution by module kind | src/vbaMemberCompletion.ts | tests/vbaMemberCompletion.test.ts | Module context | Verified |
| Typed local/param/module variable resolution | src/analyzer/completion/memberAccess.ts | tests/vbaMemberCompletion.test.ts | 5.2.3 / 5.3 (declarations) | Verified |
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
| Runtime hover resolution (after user/host/code-name) | src/analyzer/hover/resolveHover.ts | tests/vbaHover.test.ts | n/a | Verified |
| Runtime identifier completion (`runtime` kind) | src/analyzer/completion/identifierCompletion.ts | tests/vbaRuntime.test.ts | n/a | Verified |

Verification rule: built-in signatures must be transcribed from the Microsoft
VBA language reference or MS-VBAL, never invented. Names that collide with
intrinsic data types (`Date`, `Time`, `String`, `Error`) are deliberately
omitted to avoid type/function ambiguity in `As` positions.

---

## Addendum - Signature Help (Parameter Info)

Phase 7/9 adds the VBE call tip: when the caret is inside a call's argument
list, `src/analyzer/signature/signatureHelp.ts` returns the callee's signature
and the active-parameter index. The signature comes from one of three verified
sources (host members, user procedures, runtime built-ins); no signature is ever
invented, so an unknown callee yields no tip.

| Feature | Implementation File | Fixture | Source | Status |
|---|---|---|---|---|
| Caret-to-active-call resolution (paren + parenless) | src/analyzer/signature/signatureHelp.ts | tests/vbaSignatureHelp.test.ts | MS-VBAL 5.4.2 (call statements) | Verified |
| Host-member call signatures (`Workbooks.Open`, `Range.Offset`, ...) | src/analyzer/host/excelObjectModel.ts (`memberSignatures`) | tests/vbaSignatureHelp.test.ts | Office VBA object-model reference (learn.microsoft.com) | Verified |
| User procedure signatures (from AST) | src/analyzer/signature/signatureHelp.ts | tests/vbaSignatureHelp.test.ts | n/a (built from parsed `ProcedureNode`) | Verified |
| Runtime built-in signatures | src/analyzer/runtime/vbaRuntime.ts | tests/vbaSignatureHelp.test.ts | learn.microsoft.com/office/vba/language + MS-VBAL | Verified |

Verification rule: host-member signatures are transcribed from the Office VBA
object-model reference. Where a method has a large variadic tail (e.g.
`Application.Run` takes Arg1..Arg30) only the leading commonly-used parameters
are listed rather than inventing a synthetic `...` token.

---

## Addendum - Active Diagnostics

Phase 5 adds live semantic diagnostics computed from module text by
`src/analyzer/diagnostics/analyzeModule.ts` (rule catalogue in
`ruleMetadata.ts`). Each rule is high-confidence and cites the MS-VBAL section
it enforces; the engine is merged with the structural block-balance linter
(`src/vbaLinter.ts`, which covers the "Missing End .../unexpected terminator"
family) in `registerVbaDiagnostics`.

| Rule code | Meaning | MS-VBAL | Implementation File | Fixture | Status |
|---|---|---|---|---|---|
| `unterminated-string` | String literal with no closing quote | 3.3.4 (string literal token) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-procedure` | Two procedures share a name (Get/Let/Set excepted) | 5.3 (procedure declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-declaration` | Param/local redeclared in one procedure scope | 5.2 / 5.3 (declared names) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `duplicate-module-variable` | Module-level variable redeclared | 5.2.3 (module variable declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `const-assignment` | Assignment to a declared `Const` | 5.4.3.1 (Const cannot be assigned) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `option-explicit-missing` | Code module omits `Option Explicit` (configurable) | 5.2.4.1.1 (Option Explicit) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `unknown-call` | Call statement whose callee is a bare (non-member) identifier - lone identifier, parenless args (`MsgBox "hi"`), or `Call Foo` - that resolves to no project procedure, runtime function, host global, `Application` member, or in-scope name | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `invalid-proc-header` | A `Sub`/`Function`/`Property` header where a token other than `(` (or `As` for a `Function`/`Property Get`) follows the procedure name (e.g. `Sub My Sub`) | 5.3.1 (procedure declarations) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `unbalanced-parens` | A `(` left open at a statement boundary, or a `)` with no matching `(`, within one logical statement | 3.3.1 (special tokens) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `argument-count` | A call statement to a same-module, unique exported project, or module-qualified exported standard-module Sub/Function supplies too few/too many arguments (Optional/ParamArray aware), or a named argument names no parameter | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `argument-type-mismatch` | A same-module, unique exported project, module-qualified exported standard-module, or curated runtime call receives an argument whose inferred type is a provable deterministic runtime type error; focused oracle cases compile successfully but runtime probes raise error 13 for nonnumeric string-to-numeric coercion while numeric-string controls run | 5.3.1 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `argument-object-type-mismatch` | A same-module or curated runtime call receives a scalar argument where an object parameter is required; error severity because a focused VBE oracle case rejects it at compile time | 5.3.1 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `assignment-type-mismatch` | A scalar assignment, including a source-backed writable workbook class property, receives a value whose inferred type is a provable deterministic runtime type error; focused oracle cases compile successfully but runtime probes raise error 13 for nonnumeric string-to-numeric/Boolean/property coercion while valid coercion controls run | 5.4.3 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `readonly-member-assignment` | A source-backed workbook class property assignment targets a member whose source surface has no setter; focused VBE oracle evidence rejects this as `Can't assign to read-only property` | 5.4.3 / VBE oracle read-only property | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `member-not-found` | A source-backed workbook class receiver or dump-backed exhaustive host receiver uses a member name absent from the known member surface; focused VBE oracle evidence rejects unknown class property/method cases and `ThisWorkbook.DoesntExist` as `Method or data member not found`, while known member controls compile | 5.6.9 / VBE oracle member binding | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `string-arithmetic-coercion` | A numeric context contains an arithmetic expression with a provably nonnumeric string literal; error severity because focused VBE oracle cases show it compiles but deterministically raises runtime error 13 | 5.6 / runtime type coercion | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `dim-initializer` | A variable declaration includes a VB.NET-style inline initializer (`Dim x As Long = 1`), which VBA does not allow; `Const` is exempt | 5.2.3.1 (variable declaration) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `unexpected-declaration-token` | A declaration, parameter, or UDT field has an extra same-statement token after a complete `As` type name, such as `Dim s As String junk`; focused VBE oracle evidence rejects the representative `Dim` case as `Syntax error` | 5.2.3.1 / VBE oracle Syntax error | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `invalid-as-type-name` | A declaration uses a reserved runtime function name such as `Int` as an `As` type name | 3.3.5.2 / 5.2.3.1 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `set-requires-object` | A `Set` assignment targets a known intrinsic scalar variable | 5.4.3 | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `scalar-member-access` | A member-access dot targets a variable declared as a known intrinsic scalar (`String`, numeric, `Boolean`, `Date`); named scalar members are VBE Compile `Invalid qualifier` errors and trailing scalar dots are VBE Compile `Syntax error`s | 5.6.9 / VBE oracle Invalid qualifier and Syntax error | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts + syntax_corpus/oracle/vbe_oracle_cases.json | Verified |
| `call-requires-parens` | A `Call` statement supplies arguments without enclosing parentheses (`Call MsgBox "hi"`) | 5.4.2.1 (call statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `required-param-after-optional` | A required parameter follows an `Optional` parameter in a procedure header | 5.3.1.5 (parameter list) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `paramarray-not-last` | A `ParamArray` parameter is not the final parameter | 5.3.1.6 (ParamArray) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `exit-wrong-proc` | An `Exit Sub`/`Exit Function`/`Exit Property` does not match the enclosing procedure kind (`Exit Do`/`Exit For` excluded) | 5.4.1.3 (Exit statement) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |
| `option-after-declaration` | An `Option` statement appears after a declaration or procedure (only `Attribute` lines may precede it) | 5.2.1 (module options) | src/analyzer/diagnostics/analyzeModule.ts | tests/vbaDiagnostics.test.ts | Verified |

Deliberately deferred (not shipped): `undeclared-variable` (variable used
without declaration under Option Explicit) and the broad arbitrary-expression
form of `unknown-call`. The `unknown-call` rule now ships for the three
unambiguous call forms - a lone identifier, a parenless call with arguments
(`msrbox ""`), and an explicit `Call` - while the implicit-host-member form
`Cells(1, 1)` / `Range("A1")` and any statement containing a top-level `=`
(assignment) are excluded. Argument-count validation (`argument-count`) is
likewise limited to same-module and deterministic project signature calls:
ambiguous bare exported project names stay silent, while module-qualified
standard-module calls resolve through the named module only. Argument and
assignment type diagnostics ship only where the local expression model can infer
both sides deterministically and the behavior is backed by compile/runtime
oracle evidence; host arity and broad object/member type checking stay
deferred. The
remaining deferred cases require a full
expression binder plus a complete host catalogue; without them they would emit
false positives, which the project's no-false-positive rule forbids. They will
ship only once they can be proven safe. "Invalid line continuation" is also
deferred for the same reason.

Verification rule: a new diagnostic rule must (1) carry an MS-VBAL
`specReference` in `ruleMetadata.ts`, (2) be high-confidence, and (3) have
positive and negative fixtures in `tests/vbaDiagnostics.test.ts`.
