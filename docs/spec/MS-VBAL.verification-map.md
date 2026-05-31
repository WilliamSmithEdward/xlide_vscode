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
| Excel object-model member metadata | src/analyzer/host/excelObjectModel.ts | tests/vbaMemberCompletion.test.ts | Office VBA object-model reference (learn.microsoft.com), verified 2026-05-30 | Verified |
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

Deliberately deferred (not shipped): `undeclared-variable` (variable used
without declaration under Option Explicit) and `unknown-call` (unknown
procedure call). Both require a full expression binder plus a complete host
catalogue; without them they would emit false positives, which the project's
no-false-positive rule forbids. They will ship only once they can be proven
safe. "Invalid line continuation" is also deferred for the same reason.

Verification rule: a new diagnostic rule must (1) carry an MS-VBAL
`specReference` in `ruleMetadata.ts`, (2) be high-confidence, and (3) have
positive and negative fixtures in `tests/vbaDiagnostics.test.ts`.
