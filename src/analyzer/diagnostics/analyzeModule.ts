// Active VBA diagnostics engine (MS-VBAL Phase 5).
//
// `analyzeModule` runs the high-confidence semantic rules from the rule
// catalogue over one module's source and returns offset-based diagnostics. It
// is pure (no `vscode`): the editor layer converts spans to ranges and severity
// names to the VS Code enum. Structural block-balance checking stays in
// `src/vbaLinter.ts` (lintVbaSource); this engine adds the semantic rules on
// top, so the two do not overlap or double-report.
//
// Design rule (see /memories): no "looks like" heuristics. Every rule here is
// deterministic - it flags a construct only when the language guarantees it is
// an error. The one cross-module rule, `unknown-call`, fires only on a call
// statement whose callee is a bare (non-member) identifier - the lone-identifier
// form, the parenless-argument form (`MsgBox "hi"`), or `Call name` - whose name
// resolves to no procedure anywhere in the project, no VBA runtime
// function/statement, no host global or Application member, and no in-scope
// declaration - the unambiguous VBE "Sub or Function not defined" error. Broader
// flow-sensitive rules (undeclared variable, arbitrary-expression unknown call)
// still need a full expression binder and remain intentionally omitted to avoid
// false positives.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { getHostMembers, resolveHostGlobal } from '../host/hostModel';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import { STATEMENT_KEYWORDS } from '../signature/signatureHelp';
import type {
	BodyNode,
	ModuleNode,
	ProcedureNode,
	Span,
	StatementNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type { ModuleSymbolKind, VbaSymbol } from '../symbols/symbolModel';
import { isProcedureKind } from '../symbols/symbolModel';
import {
	DIAGNOSTIC_RULES,
	DiagnosticRuleName,
	DiagnosticSeverity,
} from './ruleMetadata';

/** A single diagnostic produced by the analyzer (offset-based). */
export interface VbaDiagnostic {
	/** Stable rule code (see DIAGNOSTIC_RULES). */
	code: string;
	/** Human-readable message. */
	message: string;
	/** Effective severity (after any user override). */
	severity: DiagnosticSeverity;
	/** Source span in UTF-16 offsets. */
	span: Span;
	/** MS-VBAL (or other) reference for the rule, when known. */
	specReference?: string;
}

/** Per-rule severity overrides; `'off'` disables a rule. */
export type SeverityOverrides = Partial<
	Record<DiagnosticRuleName, DiagnosticSeverity | 'off'>
>;

/** Inputs for {@link analyzeModule}. */
export interface AnalyzeModuleOptions {
	/** VB component name (used only for symbol container labels). */
	moduleName?: string;
	/** Workbook-project role of the module. */
	moduleKind?: ModuleSymbolKind;
	/** Optional per-rule severity overrides (e.g. Option Explicit severity). */
	severities?: SeverityOverrides;
	/**
	 * Lowercased names of every procedure across the whole project (from the
	 * ProjectIndex). Required for the unknown-call statement rule; when omitted,
	 * that cross-module rule does not run (so single-module analysis never
	 * false-positives on a call to a procedure in another module).
	 */
	knownProcedures?: ReadonlySet<string>;
}

/** Counts double-quote characters; an odd count means the string is unterminated. */
function countQuotes(text: string): number {
	let n = 0;
	for (const ch of text) {
		if (ch === '"') {
			n++;
		}
	}
	return n;
}

/** Resolves the effective severity of a rule, or undefined when switched off. */
function severityOf(
	rule: DiagnosticRuleName,
	overrides: SeverityOverrides | undefined,
): DiagnosticSeverity | undefined {
	const override = overrides?.[rule];
	if (override === 'off') {
		return undefined;
	}
	return override ?? DIAGNOSTIC_RULES[rule].defaultSeverity;
}

/**
 * Analyzes one VBA module source and returns its active diagnostics.
 * Never throws: any internal failure yields an empty list.
 */
export function analyzeModule(
	source: string,
	opts: AnalyzeModuleOptions = {},
): VbaDiagnostic[] {
	try {
		return runRules(source, opts);
	} catch {
		return [];
	}
}

function runRules(
	source: string,
	opts: AnalyzeModuleOptions,
): VbaDiagnostic[] {
	const out: VbaDiagnostic[] = [];
	const moduleName = opts.moduleName ?? 'Module';
	const moduleKind = opts.moduleKind ?? 'standard';
	const overrides = opts.severities;

	const push = (
		rule: DiagnosticRuleName,
		message: string,
		span: Span,
	): void => {
		const severity = severityOf(rule, overrides);
		if (!severity) {
			return;
		}
		const meta = DIAGNOSTIC_RULES[rule];
		out.push({
			code: meta.code,
			message,
			severity,
			span,
			specReference: meta.specReference,
		});
	};

	const mod = parseModule(source);
	const symbols = buildModuleSymbols(moduleName, moduleKind, source);

	checkUnterminatedStrings(source, push);
	checkDuplicateProcedures(symbols.root.children ?? [], push);
	checkDuplicateDeclarations(symbols.root.children ?? [], push);
	checkDuplicateModuleMembers(symbols.root.children ?? [], push);
	checkConstAssignment(source, mod, symbols, push);
	checkOptionExplicit(source, mod, push);
	checkProcedureHeader(source, mod, push);
	checkUnbalancedParens(source, push);
	checkArgumentCount(source, mod, push);
	if (opts.knownProcedures) {
		checkUnknownCallStatement(source, mod, symbols, opts.knownProcedures, push);
	}

	return out;
}

type PushFn = (rule: DiagnosticRuleName, message: string, span: Span) => void;

/** Rule: a string literal with an odd number of quotes is never closed. */
function checkUnterminatedStrings(source: string, push: PushFn): void {
	for (const tok of tokenize(source)) {
		if (tok.kind === 'stringLiteral' && countQuotes(tok.rawText) % 2 === 1) {
			push(
				'unterminatedString',
				'Unterminated string literal.',
				{ start: tok.start, end: tok.end },
			);
		}
	}
}

/**
 * Rule: a procedure name may name at most one Sub/Function, OR a set of distinct
 * Property accessors (one Get, one Let, one Set). Any other repeat is the VBA
 * "Ambiguous name detected" compile error.
 */
function checkDuplicateProcedures(members: VbaSymbol[], push: PushFn): void {
	const groups = new Map<string, VbaSymbol[]>();
	for (const sym of members) {
		if (!isProcedureKind(sym.kind)) {
			continue;
		}
		const key = sym.name.toLowerCase();
		(groups.get(key) ?? groups.set(key, []).get(key)!).push(sym);
	}

	for (const group of groups.values()) {
		if (group.length < 2) {
			continue;
		}
		let valueProcSeen = false;
		const accessorSeen = new Set<string>();
		for (const sym of group) {
			const isProperty =
				sym.kind === 'propertyGet' ||
				sym.kind === 'propertyLet' ||
				sym.kind === 'propertySet';
			let conflict = false;
			if (!isProperty) {
				conflict = valueProcSeen || accessorSeen.size > 0;
				valueProcSeen = true;
			} else {
				conflict = valueProcSeen || accessorSeen.has(sym.kind);
				accessorSeen.add(sym.kind);
			}
			if (conflict) {
				push(
					'duplicateProcedure',
					`Ambiguous name detected: '${sym.name}' is already declared in this module.`,
					sym.nameSpan,
				);
			}
		}
	}
}

/**
 * Rule: within one procedure, a name may be declared once across its parameters,
 * local Dim/Static variables, and local Const declarations. Repeats are the VBA
 * "Duplicate declaration in current scope" error. Procedure scope is flat in VBA
 * (no block scope), so locals from different branches still collide.
 */
function checkDuplicateDeclarations(members: VbaSymbol[], push: PushFn): void {
	for (const proc of members) {
		if (!isProcedureKind(proc.kind)) {
			continue;
		}
		const seen = new Set<string>();
		for (const child of proc.children ?? []) {
			if (
				child.kind !== 'parameter' &&
				child.kind !== 'localVariable' &&
				child.kind !== 'constant'
			) {
				continue;
			}
			const key = child.name.toLowerCase();
			if (seen.has(key)) {
				push(
					'duplicateDeclaration',
					`Duplicate declaration in current scope: '${child.name}'.`,
					child.nameSpan,
				);
			} else {
				seen.add(key);
			}
		}
	}
}

/** Rule: a module-level variable or constant declared more than once. */
function checkDuplicateModuleMembers(members: VbaSymbol[], push: PushFn): void {
	const seen = new Set<string>();
	for (const sym of members) {
		if (sym.kind !== 'moduleVariable' && sym.kind !== 'constant') {
			continue;
		}
		const key = sym.name.toLowerCase();
		if (seen.has(key)) {
			push(
				'duplicateModuleMember',
				`Duplicate declaration: '${sym.name}' is already declared at module level.`,
				sym.nameSpan,
			);
		} else {
			seen.add(key);
		}
	}
}

/**
 * Rule: assigning to a constant is illegal. High-confidence form only - the
 * left-hand side must be a bare identifier (no member access, no index) that
 * resolves to a Const declared at module level or in the enclosing procedure.
 */
function checkConstAssignment(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
): void {
	const moduleConsts = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		if (sym.kind === 'constant') {
			moduleConsts.add(sym.name.toLowerCase());
		}
	}

	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(s) => isProcedureKind(s.kind) && s.fullSpan.start === member.span.start,
		);
		const localConsts = new Set<string>();
		for (const child of procSym?.children ?? []) {
			if (child.kind === 'constant') {
				localConsts.add(child.name.toLowerCase());
			}
		}
		const inScope = (lower: string): boolean =>
			localConsts.has(lower) || moduleConsts.has(lower);
		forEachStatement(member.body, (stmt) => {
			const hit = bareAssignmentTarget(source, stmt.span);
			if (hit && inScope(hit.name.toLowerCase())) {
				push(
					'constAssignment',
					`Cannot assign to constant '${hit.name}'.`,
					hit.span,
				);
			}
		});
	}
}

/** Walks every StatementNode in a body, descending into nested blocks. */
function forEachStatement(
	body: BodyNode[],
	visit: (stmt: StatementNode) => void,
): void {
	for (const node of body) {
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit);
		}
	}
}

/**
 * If the statement spanning `span` is a simple assignment to a bare identifier
 * (`name = ...` or `Let name = ...`), returns that identifier and its span;
 * otherwise undefined. `Set` (object) assignments and any left-hand side with a
 * `.` or `(` are excluded so only true scalar-name assignments are considered.
 */
function bareAssignmentTarget(
	source: string,
	span: Span,
): { name: string; span: Span } | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	let i = 0;
	// Skip a leading line label: `Name:`.
	if (
		toks.length >= 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		i = 2;
	}
	// Skip an explicit `Let`; bail on `Set` (object assignment).
	if (toks[i] && toks[i].kind === 'keyword') {
		const kw = toks[i].rawText.toLowerCase();
		if (kw === 'set') {
			return undefined;
		}
		if (kw === 'let') {
			i++;
		}
	}
	const nameTok = toks[i];
	if (!nameTok || nameTok.kind !== 'identifier') {
		return undefined; // first token must be a plain identifier LHS
	}
	const next = toks[i + 1];
	if (!next || next.kind !== 'operator' || next.rawText !== '=') {
		return undefined; // not `name =` (excludes `.`, `(`, `<=`, `<>`, comparisons)
	}
	return {
		name: nameTok.rawText,
		span: { start: span.start + nameTok.start, end: span.start + nameTok.end },
	};
}

/**
 * Rule: a *call statement* whose callee is a bare (non-member) identifier - the
 * lone-identifier form `DoStartup`, the parenless-argument form `MsgBox "hi"` /
 * `Foo 1, 2`, or the explicit `Call DoWork` / `Call Foo(1, 2)` form - is a call
 * to a Sub/Function of that name. When the name resolves to nothing the VBE
 * raises "Sub or Function not defined".
 *
 * A name is considered resolved when it matches any project procedure, a name
 * declared in the current module (procedures, module variables/consts, types,
 * enums and their members, Declares), a parameter/local/const of the enclosing
 * procedure, a VBA runtime function/statement, or a host global / Application
 * member (Excel exposes Application's members in the global scope). The callee
 * detection ({@link callStatementTarget}) deliberately ignores assignments,
 * member calls, line labels, and the bare `Name(...)` indexed/implicit-member
 * form so those never produce a false positive.
 */
function checkUnknownCallStatement(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	knownProcedures: ReadonlySet<string>,
	push: PushFn,
): void {
	// Names visible module-wide: every module-level declaration (including a
	// live, not-yet-saved procedure) plus enum member names.
	const moduleNames = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		moduleNames.add(sym.name.toLowerCase());
		if (sym.kind === 'enum') {
			for (const member of sym.children ?? []) {
				moduleNames.add(member.name.toLowerCase());
			}
		}
	}

	// Excel injects Application's members into the global scope, so a bare call
	// may legitimately bind to one of them (Calculate, Volatile, Evaluate, ...).
	const appType = resolveHostGlobal('Application');
	const appMembers = new Set(
		(appType ? getHostMembers(appType) : []).map((mm) => mm.name.toLowerCase()),
	);

	const isKnown = (name: string, locals: ReadonlySet<string>): boolean => {
		const lower = name.toLowerCase();
		return (
			knownProcedures.has(lower) ||
			moduleNames.has(lower) ||
			locals.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(s) => isProcedureKind(s.kind) && s.fullSpan.start === member.span.start,
		);
		const locals = new Set<string>();
		for (const child of procSym?.children ?? []) {
			locals.add(child.name.toLowerCase());
		}
		forEachStatement(member.body, (stmt) => {
			const hit = callStatementTarget(source, stmt.span);
			if (hit && !isKnown(hit.name, locals)) {
				push(
					'unknownCallStatement',
					`Sub or Function not defined: '${hit.name}'.`,
					hit.span,
				);
			}
		});
	}
}

/**
 * If the statement spanning `span` is a *call statement* whose callee is a bare
 * (non-member) identifier, returns that identifier and its absolute span;
 * otherwise undefined. Three forms qualify:
 *
 *   - a lone identifier (`DoStartup`),
 *   - a parenless call with arguments (`MsgBox "hi"`, `Foo 1, 2`), and
 *   - an explicit `Call` statement (`Call DoWork`, `Call Foo(1, 2)`).
 *
 * Disqualified (returns undefined): assignments (a top-level `=`), member calls
 * (a leading `receiver.`), line labels (`done:`), a statement keyword leader
 * (`Set`, `Open`, `Print`, ...), and the parenthesized/indexed form
 * `Name(...)...` for a non-`Call` statement - the latter is excluded because a
 * bare `Cells(1, 1)` or `Range("A1")` is an implicit Application member and must
 * never be flagged. `Call Name(...)` is still inspected because `Call` makes the
 * call unambiguous.
 */
function callStatementTarget(
	source: string,
	span: Span,
): { name: string; span: Span } | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0) {
		return undefined;
	}

	let idx = 0;
	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
	if (explicitCall) {
		idx = 1;
	}

	const callee = toks[idx];
	if (!callee || callee.kind !== 'identifier') {
		return undefined;
	}
	if (STATEMENT_KEYWORDS.has(callee.rawText.toLowerCase())) {
		return undefined;
	}

	const result = {
		name: callee.rawText,
		span: { start: span.start + callee.start, end: span.start + callee.end },
	};

	const next = toks[idx + 1];
	if (!next) {
		// A lone identifier (optionally `Call name`). For the non-Call form, an
		// identifier immediately followed by `:` is a line label, not a call; the
		// statement span excludes the colon, so peek past it in the full source.
		if (!explicitCall) {
			let j = span.start + callee.end;
			while (j < source.length && (source[j] === ' ' || source[j] === '\t')) {
				j++;
			}
			if (source[j] === ':') {
				return undefined;
			}
		}
		return result;
	}

	const r = next.rawText;
	if (r === '.' || r === ':') {
		return undefined; // member call or label/separator artifact
	}
	if (!explicitCall && r === '(') {
		return undefined; // index/group/implicit host member - defer
	}
	if (!explicitCall) {
		// The first argument must be separated from the callee by whitespace; this
		// is what makes it a parenless call rather than e.g. `a=b` glued together.
		const gap = source.slice(span.start + callee.end, span.start + next.start);
		if (!/\s/.test(gap)) {
			return undefined;
		}
	}
	// A top-level `=` (outside any parentheses/brackets) makes this an assignment.
	let depth = 0;
	for (let k = idx + 1; k < toks.length; k++) {
		const tr = toks[k].rawText;
		if (tr === '(' || tr === '[') {
			depth++;
		} else if (tr === ')' || tr === ']') {
			depth--;
		} else if (depth === 0 && tr === '=') {
			return undefined;
		}
	}
	return result;
}

/** Access/storage modifiers that may lead a procedure declaration. */
const PROC_MODIFIERS = new Set([
	'public', 'private', 'friend', 'global', 'static',
]);

/**
 * Rule: a procedure header must be `[(modifiers)] Sub|Function|Property Get/Let/Set
 * Name [(params)] [As Type]`. Once the name is read, the only legal next token is
 * `(` (the parameter list) or, for a `Function`/`Property Get`, `As` (the return
 * type). Any other token - most commonly a second word, as in `Sub My Sub`, where
 * the name was meant to contain a space - is the VBE "Expected: (" compile error.
 * Property `Let`/`Set` and `Sub` have no return value, so an `As` right after the
 * name is rejected for them too.
 */
function checkProcedureHeader(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const headerStart = member.span.start;
		const nl = source.indexOf('\n', headerStart);
		const headerEnd = nl === -1 ? member.span.end : nl;
		const toks = tokenize(source.slice(headerStart, headerEnd)).filter(
			(t) => t.kind !== 'comment' && t.kind !== 'newline',
		);

		let i = 0;
		while (i < toks.length && PROC_MODIFIERS.has(toks[i].rawText.toLowerCase())) {
			i++;
		}
		const kw = toks[i]?.rawText.toLowerCase();
		let allowAs = false;
		if (kw === 'function') {
			allowAs = true;
			i++;
		} else if (kw === 'sub') {
			i++;
		} else if (kw === 'property') {
			i++;
			if (toks[i]?.rawText.toLowerCase() === 'get') {
				allowAs = true;
			}
			i++; // skip the accessor (Get/Let/Set)
		} else {
			continue; // not a recognised procedure header
		}

		const nameTok = toks[i];
		if (!nameTok) {
			continue; // malformed in a way the structural linter already reports
		}
		const next = toks[i + 1];
		if (!next) {
			continue; // `Sub Foo` with no parameter list is legal
		}
		const r = next.rawText;
		if (r === '(' || (allowAs && r.toLowerCase() === 'as')) {
			continue;
		}
		push(
			'invalidProcedureHeader',
			`Unexpected '${r}' after procedure name '${stripHeaderBrackets(nameTok.rawText)}'; a procedure name must be a single identifier.`,
			{ start: headerStart + next.start, end: headerStart + next.end },
		);
	}
}

/** Strips the surrounding `[ ]` from a bracketed identifier, if present. */
function stripHeaderBrackets(text: string): string {
	return text.startsWith('[') && text.endsWith(']')
		? text.slice(1, -1)
		: text;
}

/**
 * Rule: every parenthesis must be matched within its logical statement. VBA has
 * no cross-statement parentheses (a `(` is closed before the line ends unless a
 * `_` line-continuation joins the next physical line, which the lexer already
 * folds into trivia), so an open `(` left dangling at a statement boundary, or a
 * `)` with no matching `(`, is always the VBE "Expected: )" / "Syntax error".
 *
 * The scan walks the whole module's token stream, tracking paren depth and
 * resetting at each logical-statement boundary (a `newline` token or a depth-0
 * `:` statement separator). Only literal `(`/`)` punctuation tokens count -
 * parentheses inside strings, comments, date literals, and `[bracketed]` names
 * are distinct token kinds, so they can never create a false positive. At most
 * one diagnostic is reported per statement.
 */
function checkUnbalancedParens(source: string, push: PushFn): void {
	const toks = tokenize(source);
	let depth = 0;
	const openOffsets: number[] = [];
	let flagged = false;

	const flush = (): void => {
		if (!flagged && depth > 0) {
			const off = openOffsets[0];
			push(
				'unbalancedParens',
				"Unbalanced parentheses: a ')' is missing.",
				{ start: off, end: off + 1 },
			);
		}
		depth = 0;
		openOffsets.length = 0;
		flagged = false;
	};

	for (const tok of toks) {
		if (tok.kind === 'newline') {
			flush();
			continue;
		}
		if (tok.kind === 'colon' && depth === 0) {
			flush();
			continue;
		}
		if (tok.kind !== 'punctuation') {
			continue;
		}
		if (tok.rawText === '(') {
			depth++;
			openOffsets.push(tok.start);
		} else if (tok.rawText === ')') {
			if (depth === 0) {
				if (!flagged) {
					push(
						'unbalancedParens',
						"Unbalanced parentheses: an unexpected ')' was found.",
						{ start: tok.start, end: tok.end },
					);
					flagged = true;
				}
			} else {
				depth--;
				openOffsets.pop();
			}
		}
	}
	flush();
}

/** A resolved call statement: callee plus its top-level argument slots. */
interface CallArguments {
	/** Callee identifier text. */
	name: string;
	/** Absolute span of the callee identifier. */
	nameSpan: Span;
	/**
	 * Top-level, comma-separated argument groups. An empty list means no
	 * arguments were supplied; an empty inner array is an omitted positional
	 * argument (`Foo 1, , 3`).
	 */
	slots: VbaToken[][];
	/** Absolute offset of the statement slice the slot tokens are relative to. */
	sliceStart: number;
}

/**
 * Rule: a call to a Sub/Function defined in *this* module must supply an
 * argument count the procedure's parameter list accepts. We validate only
 * current-module Sub/Function procedures because that is the only place we have
 * the ground-truth parameter list from the AST - host members and runtime
 * functions (partial/variadic metadata) and cross-module procedures (names only)
 * are deliberately not arity-checked, to stay false-positive-free. Property
 * accessors are skipped (they are not invoked through call-statement syntax) and
 * any name that is ambiguous in the module (duplicate/overloaded) is skipped.
 *
 * Only the unambiguous call-statement forms that {@link callStatementTarget}
 * already resolves are inspected: the parenless `Foo 1, 2` and the explicit
 * `Call Foo(1, 2)`. Expression-embedded calls (`x = Foo(1, 2)`) and the
 * parenthesized bare form `Foo(1, 2)` are not validated (the latter is
 * indistinguishable from `Cells(1, 1)` indexing without a binder).
 */
function checkArgumentCount(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	const procsByName = new Map<string, ProcedureNode[]>();
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		if (member.procKind !== 'Sub' && member.procKind !== 'Function') {
			continue;
		}
		const key = member.name.toLowerCase();
		const arr = procsByName.get(key);
		if (arr) {
			arr.push(member);
		} else {
			procsByName.set(key, [member]);
		}
	}
	if (procsByName.size === 0) {
		return;
	}

	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const call = extractCall(source, stmt.span);
			if (!call) {
				return;
			}
			const candidates = procsByName.get(call.name.toLowerCase());
			// Skip unknown names (owned by unknown-call) and ambiguous ones (e.g. a
			// duplicate definition) where the target signature is not unique.
			if (!candidates || candidates.length !== 1) {
				return;
			}
			validateArity(candidates[0], call, push);
		});
	}
}

/**
 * If the statement spanning `span` is a bare call statement, returns the callee
 * and its top-level argument slots; otherwise undefined. Reuses
 * {@link callStatementTarget} for the safe call-detection gating, then peels off
 * the argument region (the parenless tail, or the contents of the `Call`
 * statement's parentheses).
 */
function extractCall(source: string, span: Span): CallArguments | undefined {
	const hit = callStatementTarget(source, span);
	if (!hit) {
		return undefined;
	}
	const sliceStart = span.start;
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const relCalleeStart = hit.span.start - sliceStart;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0) {
		return undefined;
	}

	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
	const next = toks[calleeIdx + 1];
	let argToks: VbaToken[];
	if (explicitCall) {
		if (next && next.kind === 'punctuation' && next.rawText === '(') {
			// Collect the tokens strictly inside the call's parentheses.
			let depth = 0;
			let closed = false;
			const inner: VbaToken[] = [];
			for (let k = calleeIdx + 1; k < toks.length; k++) {
				const t = toks[k];
				if (t.kind === 'punctuation' && t.rawText === '(') {
					depth++;
					if (depth === 1) {
						continue; // skip the opening paren itself
					}
				} else if (t.kind === 'punctuation' && t.rawText === ')') {
					depth--;
					if (depth === 0) {
						closed = true;
						break;
					}
				}
				if (depth >= 1) {
					inner.push(t);
				}
			}
			if (!closed) {
				return undefined; // unbalanced - the parentheses rule reports this
			}
			argToks = inner;
		} else {
			argToks = []; // `Call Foo` with no parameter list
		}
	} else {
		argToks = toks.slice(calleeIdx + 1);
	}

	const slots = argToks.length === 0 ? [] : splitArgSlots(argToks);
	return { name: hit.name, nameSpan: hit.span, slots, sliceStart };
}

/** Splits an argument token run into top-level (depth-0) comma-separated slots. */
function splitArgSlots(toks: VbaToken[]): VbaToken[][] {
	const slots: VbaToken[][] = [[]];
	let depth = 0;
	for (const t of toks) {
		if (t.kind === 'punctuation' && t.rawText === '(') {
			depth++;
		} else if (t.kind === 'punctuation' && t.rawText === ')') {
			depth--;
		}
		if (t.kind === 'punctuation' && t.rawText === ',' && depth === 0) {
			slots.push([]);
		} else {
			slots[slots.length - 1].push(t);
		}
	}
	return slots;
}

/** True if a slot is a named argument (`name := value`). */
function isNamedSlot(slot: VbaToken[]): boolean {
	return (
		slot.length >= 2 &&
		(slot[0].kind === 'identifier' || slot[0].kind === 'bracketedIdentifier') &&
		slot[1].kind === 'operator' &&
		slot[1].rawText === ':='
	);
}

/** Describes a procedure's acceptable argument-count range for a message. */
function describeArity(required: number, max: number): string {
	if (max === Infinity) {
		return `at least ${required} argument${required === 1 ? '' : 's'}`;
	}
	if (required === max) {
		return `${required} argument${required === 1 ? '' : 's'}`;
	}
	return `between ${required} and ${max} arguments`;
}

/**
 * Validates one call's argument list against a procedure's parameters. When the
 * call uses named arguments, each name is checked against the parameter names
 * and the positional count check is skipped (positional/named mixing is too
 * subtle to count safely); otherwise the supplied slot count is checked against
 * the required minimum and the maximum implied by `Optional`/`ParamArray`.
 */
function validateArity(
	proc: ProcedureNode,
	call: CallArguments,
	push: PushFn,
): void {
	const params = proc.params;
	let required = params.length;
	for (let k = 0; k < params.length; k++) {
		if (params[k].optional || params[k].paramArray) {
			required = k;
			break;
		}
	}
	const hasParamArray = params.some((p) => p.paramArray);
	const max = hasParamArray ? Infinity : params.length;

	const named = call.slots.filter(isNamedSlot);
	if (named.length > 0) {
		const paramNames = new Set(
			params.map((p) => stripHeaderBrackets(p.name).toLowerCase()),
		);
		for (const slot of named) {
			const raw = stripHeaderBrackets(slot[0].rawText);
			if (!paramNames.has(raw.toLowerCase())) {
				push(
					'argumentCount',
					`Named argument not found: '${raw}' is not a parameter of '${proc.name}'.`,
					{
						start: call.sliceStart + slot[0].start,
						end: call.sliceStart + slot[0].end,
					},
				);
			}
		}
		return; // positional count is not validated alongside named arguments
	}

	const n = call.slots.length;
	if (n < required || n > max) {
		push(
			'argumentCount',
			`Wrong number of arguments to '${proc.name}': expected ${describeArity(required, max)}, but got ${n}.`,
			call.nameSpan,
		);
	}
}

/**
 * Rule: a code module that contains real code but no `Option Explicit` lets
 * variables be used without declaration. Empty/attribute-only modules are
 * skipped to avoid noise on blank document modules.
 */
function checkOptionExplicit(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	let hasExplicit = false;
	let hasCode = false;
	for (const member of mod.members) {
		if (member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim())) {
			hasExplicit = true;
		}
		if (
			member.kind === 'Procedure' ||
			member.kind === 'VariableGroup' ||
			member.kind === 'Type' ||
			member.kind === 'Enum' ||
			member.kind === 'Declare'
		) {
			hasCode = true;
		}
	}
	if (hasExplicit || !hasCode) {
		return;
	}
	// Point at the first physical line so the squiggle sits at the module top.
	const nl = source.search(/\r|\n/);
	const end = nl === -1 ? source.length : nl;
	push(
		'optionExplicitMissing',
		'Option Explicit is not specified; variables can be used without being declared. Add "Option Explicit" to the top of the module.',
		{ start: 0, end },
	);
}
