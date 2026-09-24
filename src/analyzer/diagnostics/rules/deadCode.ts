// Dead code: declarations nothing reads, private procedures nothing calls,
// and statements nothing can reach.
//
// Every finding here is a structural fact about the module's own text, not a
// guess about how it runs, and each one stops at the point where the text can
// no longer prove it:
//
//   - A local or module-private variable is "unused" when no token in its
//     scope names it, and "never read" when every token that does is a plain
//     assignment to it. Passing it to a procedure counts as a read (ByRef may
//     fill it), so does indexing it, a member access on it, a `For` loop over
//     it, and any mention inside an inactive `#If` arm.
//   - Only Private procedures are reported as uncalled. A Public one may be
//     wired to a button, a shape, the ribbon, a hotkey, `OnTime` or
//     `Application.Run` in a file XLIDE cannot see, so its silence proves
//     nothing. A Private one can only be reached from its own module - or by
//     name in a string, which is why every string literal in the project is
//     searched too. Event handlers and procedures with member attributes are
//     never reported: the host calls those.
//   - Code after `Exit Sub`, `Exit Function`, `Exit Property`, `Exit Do`,
//     `Exit For`, `GoTo`, `Resume`, `Return` or `End` in the same block is
//     unreachable until a label, a line number, a `Case` or a `#If` gives the
//     flow somewhere to land.

import type { VbaToken } from '../../lexer/tokenKinds';
import { tokenizeCached } from '../../lexer/tokenize';
import { identifierWords } from '../../lexer/tokenHelpers';
import type {
	BodyNode,
	LeafStatementNode,
	ModuleNode,
	ProcedureNode,
	Span,
	VariableGroupNode,
} from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { classifyReferenceKinds } from '../../references/referenceKinds';
import type { ModuleSymbolKind } from '../../symbols/symbolModel';
import { isProcedureKind } from '../../symbols/symbolModel';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { eventHandlerProcedureForName } from '../../completion/eventHandlers';
import type { PushFn } from '../analysisContext';
import { statementTokens } from '../analysisContext';
import {
	activeModuleMembers,
	declaredNameSpan,
	forEachVariableGroup,
	isInactiveNode,
	tokenText,
} from '../walker';
import { wholeLineSpan } from '../../../vbaSourceScan';
import { attachedCommentsStart } from '../../docs/docComment';

interface TrackedDeclaration {
	name: string;
	lower: string;
	nameSpan: Span;
	group: VariableGroupNode;
	scope: 'local' | 'module';
}

interface Reference {
	offset: number;
	/**
	 * Read by construction: a `For`/`For Each` control variable is read by
	 * the loop itself, and `x(i) = v` reads the array it stores into.
	 */
	readByForm: boolean;
}

// ------------------------------------------------------------ unused names

/**
 * Locals nothing uses, module-private variables and constants nothing uses,
 * and variables whose only mentions assign to them.
 */
export function checkUnusedDeclarations(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const tokens = tokenizeCached(source);
	const attributed = new Set<number>();
	for (const symbol of symbols.all) {
		if (symbol.attributes && symbol.attributes.length > 0) {
			attributed.add(symbol.nameSpan.start);
		}
	}

	const members = activeModuleMembers(mod, activity);
	const procedures = members.filter((member): member is ProcedureNode => member.kind === 'Procedure');

	// Module-level declarations that only this module can name.
	const moduleLevel: TrackedDeclaration[] = [];
	for (const member of members) {
		if (member.kind !== 'VariableGroup' || member.withEvents || !isModulePrivate(member)) {
			continue;
		}
		for (const decl of member.declarations) {
			if (!decl.nameSpan || attributed.has(decl.nameSpan.start)) {
				continue;
			}
			moduleLevel.push({ name: decl.name, lower: decl.name.toLowerCase(), nameSpan: decl.nameSpan, group: member, scope: 'module' });
		}
	}

	// Procedures whose parameters or locals shadow a module-level name: a
	// mention inside them binds to the shadow, not the module variable.
	const shadowing = new Map<string, Span[]>();
	for (const proc of procedures) {
		const names = new Set<string>();
		for (const param of proc.params) {
			names.add(param.name.toLowerCase());
		}
		forEachVariableGroup(proc.body, (group) => {
			for (const decl of group.declarations) {
				names.add(decl.name.toLowerCase());
			}
		}, activity);
		for (const lower of names) {
			const spans = shadowing.get(lower) ?? [];
			spans.push(proc.span);
			shadowing.set(lower, spans);
		}
	}

	// One pass over the token stream serves every scope: the module-level
	// names see the whole module, each procedure's locals see its own span.
	// Reference kinds are then classified in one walk as well, so the cost
	// stays linear in module size however many procedures it has.
	const scopes: Array<{ from: number; to: number; declarations: TrackedDeclaration[] }> = [];
	if (moduleLevel.length > 0) {
		scopes.push({ from: 0, to: source.length, declarations: moduleLevel });
	}
	for (const proc of procedures) {
		const locals: TrackedDeclaration[] = [];
		forEachVariableGroup(proc.body, (group) => {
			for (const decl of group.declarations) {
				if (!decl.nameSpan || attributed.has(decl.nameSpan.start)) {
					continue;
				}
				locals.push({ name: decl.name, lower: decl.name.toLowerCase(), nameSpan: decl.nameSpan, group, scope: 'local' });
			}
		}, activity);
		if (locals.length > 0) {
			scopes.push({ from: proc.span.start, to: proc.span.end, declarations: locals });
		}
	}
	if (scopes.length === 0) {
		return;
	}

	const referencesByScope = scopes.map((scope) => collectReferences(
		tokens,
		scope.from,
		scope.to,
		new Set(scope.declarations.map((decl) => decl.nameSpan.start)),
	));
	const allOffsets: number[] = [];
	for (const references of referencesByScope) {
		for (const list of references.values()) {
			for (const ref of list) {
				allOffsets.push(ref.offset);
			}
		}
	}
	const kinds = classifyReferenceKinds(source, allOffsets);

	scopes.forEach((scope, index) => {
		const references = referencesByScope[index];
		for (const decl of scope.declarations) {
			let own = references.get(decl.lower) ?? [];
			if (decl.scope === 'module') {
				const shadows = shadowing.get(decl.lower) ?? [];
				own = own.filter((ref) => !shadows.some((span) => ref.offset >= span.start && ref.offset < span.end));
			}
			report(source, decl, own, kinds, push);
		}
	});
}

function isModulePrivate(group: VariableGroupNode): boolean {
	const modifier = group.modifier.toLowerCase();
	if (modifier === 'private' || modifier === 'dim') {
		return true;
	}
	// A bare `Const` at module level is private by default (MS-VBAL 5.2.3.2).
	return group.isConst && modifier === '';
}

function report(
	source: string,
	decl: TrackedDeclaration,
	references: Reference[],
	kinds: ReadonlyMap<number, 'read' | 'write' | 'readwrite'>,
	push: PushFn,
): void {
	if (references.length === 0) {
		const what = decl.group.isConst
			? 'Constant'
			: decl.scope === 'module' ? 'Module-level variable' : 'Local variable';
		push(
			'unusedVariable',
			`${what} '${decl.name}' is declared but never used.`,
			decl.nameSpan,
			{ removeDeclaration: removeDeclarationData(source, decl) },
		);
		return;
	}
	if (decl.group.isConst) {
		return;
	}
	const everWritten = references.some((ref) => kinds.get(ref.offset) === 'write' && !ref.readByForm);
	const everRead = references.some((ref) => ref.readByForm || kinds.get(ref.offset) !== 'write');
	if (everWritten && !everRead) {
		push(
			'variableNeverRead',
			`Variable '${decl.name}' is assigned but its value is never read.`,
			decl.nameSpan,
		);
	}
}

/**
 * Every mention of a name within [from, to) that could bind to a variable:
 * not a member name, not a named-argument name, not a declaration site.
 * Keyed by lowercased name.
 */
function collectReferences(
	tokens: readonly VbaToken[],
	from: number,
	to: number,
	declared: ReadonlySet<number>,
): Map<string, Reference[]> {
	const out = new Map<string, Reference[]>();
	const first = firstTokenAtOrAfter(tokens, from);
	let prev: VbaToken | undefined = tokens[first - 1];
	let prev2: VbaToken | undefined = tokens[first - 2];
	for (let i = first; i < tokens.length; i++) {
		const token = tokens[i];
		if (token.start >= to) {
			break;
		}
		const isName = token.kind === 'identifier' || token.kind === 'keyword';
		if (isName && !declared.has(token.start) && !isMemberName(prev) && !isNamedArgument(tokens[i + 1])) {
			const lower = tokenText(token);
			const next = tokens[i + 1];
			const readByForm = tokenText(prev) === 'for'
				|| (tokenText(prev) === 'each' && tokenText(prev2) === 'for')
				|| (!!next && next.kind === 'punctuation' && next.rawText === '(');
			const list = out.get(lower) ?? [];
			list.push({ offset: token.start, readByForm });
			out.set(lower, list);
		}
		if (token.kind !== 'comment') {
			prev2 = prev;
			prev = token;
		}
	}
	return out;
}

/** Index of the first token starting at or after `offset` (tokens are in source order). */
function firstTokenAtOrAfter(tokens: readonly VbaToken[], offset: number): number {
	let low = 0;
	let high = tokens.length;
	while (low < high) {
		const mid = (low + high) >> 1;
		if (tokens[mid].start < offset) {
			low = mid + 1;
		} else {
			high = mid;
		}
	}
	return low;
}

function isMemberName(prev: VbaToken | undefined): boolean {
	return !!prev && ((prev.kind === 'punctuation' && prev.rawText === '.')
		|| (prev.kind === 'operator' && prev.rawText === '!'));
}

function isNamedArgument(next: VbaToken | undefined): boolean {
	return !!next && next.kind === 'operator' && next.rawText === ':=';
}

/**
 * The edit that removes one declaration: its whole line when it stands
 * alone there, else its own name (and separator) from a `Dim a, b` list.
 * Absent when the statement shares its line with something else.
 */
function removeDeclarationData(
	source: string,
	decl: TrackedDeclaration,
): { variableName: string; edit: { span: Span; newText: string } } | undefined {
	const group = decl.group;
	const { start: lineStart, end: lineEnd } = wholeLineSpan(source, group.span);
	const lineText = source.slice(lineStart, lineEnd);
	if (group.declarations.length === 1) {
		// Only the declaration (and perhaps a comment) on its line.
		const statementText = source.slice(group.span.start, group.span.end);
		const rest = lineText.replace(statementText, '').replace(/\r?\n$/, '').trim();
		if (rest !== '' && !rest.startsWith("'")) {
			return undefined;
		}
		if (source.slice(group.span.start, group.span.end).includes('\n')) {
			return undefined;
		}
		// A module variable's doc comment goes with it; left behind, it would
		// document the declaration that came next.
		const start = decl.scope === 'module' ? attachedCommentsStart(source, group.span.start) : lineStart;
		return { variableName: decl.name, edit: { span: { start, end: lineEnd }, newText: '' } };
	}
	const index = group.declarations.findIndex((d) => d.nameSpan?.start === decl.nameSpan.start);
	if (index < 0) {
		return undefined;
	}
	const own = group.declarations[index];
	if (index < group.declarations.length - 1) {
		const next = group.declarations[index + 1];
		return { variableName: decl.name, edit: { span: { start: own.span.start, end: next.span.start }, newText: '' } };
	}
	const previous = group.declarations[index - 1];
	return { variableName: decl.name, edit: { span: { start: previous.span.end, end: own.span.end }, newText: '' } };
}

// ---------------------------------------------------- uncalled procedures

const AUTO_MACRO = /^auto_(open|close|activate|deactivate|exec|new|add|remove)$/i;

/** Private procedures no token in the module and no string in the project names. */
export function checkUnusedPrivateProcedures(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	projectStringLiteralWords: ReadonlySet<string> | undefined,
	push: PushFn,
): void {
	const members = activeModuleMembers(mod, activity);
	const candidates = members.filter((member): member is ProcedureNode =>
		member.kind === 'Procedure'
		&& member.modifiers.some((modifier) => modifier.toLowerCase() === 'private')
		&& !!member.nameSpan
		&& !hasMemberAttribute(source, member, symbols)
		&& !isHostCalled(member.name, moduleKind),
	);
	if (candidates.length === 0) {
		return;
	}
	const tokens = tokenizeCached(source);
	const declarationSites = new Set<number>();
	for (const member of mod.members) {
		if (member.kind === 'Procedure' && member.nameSpan) {
			declarationSites.add(member.nameSpan.start);
		}
	}
	// Every mention of a name, by offset: a Function assigning its own return
	// value names itself, and a procedure calling only itself is still dead,
	// so mentions inside the procedure's own body do not count for it.
	const mentions = new Map<string, number[]>();
	const ownStringWords = new Set<string>();
	const stringWords: ReadonlySet<string> = projectStringLiteralWords ?? ownStringWords;
	for (const token of tokens) {
		if ((token.kind === 'identifier' || token.kind === 'keyword') && !declarationSites.has(token.start)) {
			const lower = tokenText(token);
			const offsets = mentions.get(lower) ?? [];
			offsets.push(token.start);
			mentions.set(lower, offsets);
		} else if (token.kind === 'stringLiteral' && !projectStringLiteralWords) {
			for (const word of identifierWords(token.rawText)) {
				ownStringWords.add(word);
			}
		}
	}
	for (const proc of candidates) {
		const lower = proc.name.toLowerCase();
		const outside = (mentions.get(lower) ?? []).some((offset) => offset < proc.span.start || offset >= proc.span.end);
		if (outside || stringWords.has(lower)) {
			continue;
		}
		const kind = proc.procKind === 'Sub' ? 'Sub' : proc.procKind === 'Function' ? 'Function' : 'Property';
		push(
			'unusedProcedure',
			kind === 'Property'
				? `Private Property '${proc.name}' is never used.`
				: `Private ${kind} '${proc.name}' is never called.`,
			declaredNameSpan(source, proc.span, proc.name),
		);
	}
}

/**
 * Whether a procedure carries an `Attribute` line: a hotkey, a description,
 * a default member. An exported module writes the line inside the procedure
 * under its header; the symbol table attaches one written after it.
 */
function hasMemberAttribute(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): boolean {
	if (proc.attributes && proc.attributes.length > 0) {
		return true;
	}
	const symbol = symbols.all.find((s) => isProcedureKind(s.kind) && s.fullSpan.start === proc.span.start);
	if (symbol?.attributes && symbol.attributes.length > 0) {
		return true;
	}
	return proc.body.some((node) =>
		node.kind === 'Statement' && /^\s*Attribute\b/i.test(source.slice(node.span.start, node.span.end)),
	);
}

/**
 * Whether the host, not code, calls a procedure by this name: an event
 * handler (`Worksheet_Change`, `CommandButton1_Click`, `Class_Initialize`,
 * an interface member `IFoo_Bar`) in an object module, or an `Auto_Open`
 * style macro in a standard one.
 */
function isHostCalled(name: string, moduleKind: ModuleSymbolKind): boolean {
	if (eventHandlerProcedureForName(name)) {
		return true;
	}
	if (moduleKind === 'standard') {
		return AUTO_MACRO.test(name);
	}
	return name.includes('_');
}

// ------------------------------------------------------- unreachable code

interface Terminator {
	text: string;
}

/** Statements after an unconditional exit in the same block, until a landing point. */
export function checkUnreachableCode(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		walkBody(member.body);
	}

	function walkBody(body: BodyNode[]): void {
		let terminator: Terminator | undefined;
		let dead: Span | undefined;
		const flush = (): void => {
			if (dead && terminator) {
				push(
					'unreachableCode',
					`Unreachable code after '${terminator.text}'.`,
					dead,
					{ removeUnreachableCode: { edit: { span: wholeLineSpan(source, dead), newText: '' } } },
				);
			}
			dead = undefined;
			terminator = undefined;
		};
		for (const node of body) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'ConditionalDirective') {
				flush();
				continue;
			}
			if (isLeafStatement(node) && node.singleLineIfTail) {
				// It runs only with its single-line If's branch (MS-VBAL 5.4.2.9):
				// an Exit there ends nothing, and it is dead when its If is.
				if (terminator) {
					dead = dead ? { start: dead.start, end: node.span.end } : { start: node.span.start, end: node.span.end };
				}
				continue;
			}
			if (isLeafStatement(node)) {
				const toks = statementTokens(source, node.span);
				if (isLandingPoint(source, node, toks)) {
					flush();
					const after = tokensAfterLineNumber(toks);
					const exit = terminalStatement(after);
					if (exit) {
						terminator = exit;
					}
					continue;
				}
				if (terminator) {
					dead = dead ? { start: dead.start, end: node.span.end } : { start: node.span.start, end: node.span.end };
					continue;
				}
				const exit = terminalStatement(toks);
				if (exit) {
					terminator = exit;
				}
				continue;
			}
			// A block node.
			if (terminator) {
				if (blockHasLandingPoint(source, node)) {
					flush();
					walkBlock(node);
					continue;
				}
				dead = dead ? { start: dead.start, end: node.span.end } : { start: node.span.start, end: node.span.end };
				continue;
			}
			walkBlock(node);
		}
		flush();
	}

	function walkBlock(node: BodyNode): void {
		if (node.kind === 'IfBlock') {
			for (const branch of node.branches) {
				walkBody(branch.body);
			}
			return;
		}
		if ('body' in node && Array.isArray(node.body)) {
			walkBody(node.body);
		}
	}
}

/** A label, a line number, or a `Case` arm: somewhere control can arrive. */
function isLandingPoint(source: string, node: LeafStatementNode, toks: readonly VbaToken[]): boolean {
	if (toks.length === 0) {
		return false;
	}
	if (toks[0].kind === 'integerLiteral') {
		return true;
	}
	if (node.kind === 'Statement' && toks.length === 1
		&& (toks[0].kind === 'identifier' || toks[0].kind === 'keyword')
		&& source[node.span.end] === ':') {
		return true;
	}
	const head = tokenText(toks[0]);
	return head === 'case' || head === 'else' || head === 'elseif';
}

function tokensAfterLineNumber(toks: readonly VbaToken[]): readonly VbaToken[] {
	return toks.length > 0 && toks[0].kind === 'integerLiteral' ? toks.slice(1) : toks;
}

function terminalStatement(toks: readonly VbaToken[]): Terminator | undefined {
	if (toks.length === 0) {
		return undefined;
	}
	const head = tokenText(toks[0]);
	const second = tokenText(toks[1]);
	if (head === 'exit' && (second === 'sub' || second === 'function' || second === 'property'
		|| second === 'do' || second === 'for')) {
		return { text: `Exit ${toks[1].canonicalText ?? toks[1].rawText}` };
	}
	if (head === 'goto' && toks.length >= 2) {
		return { text: `GoTo ${toks[1].rawText}` };
	}
	if (head === 'resume') {
		return { text: toks.length === 1 ? 'Resume' : `Resume ${toks[1].canonicalText ?? toks[1].rawText}` };
	}
	if (head === 'end' && toks.length === 1) {
		return { text: 'End' };
	}
	if (head === 'return' && toks.length === 1) {
		return { text: 'Return' };
	}
	return undefined;
}

function blockHasLandingPoint(source: string, node: BodyNode): boolean {
	const bodies: BodyNode[][] = [];
	if (node.kind === 'IfBlock') {
		bodies.push(...node.branches.map((branch) => branch.body));
	} else if ('body' in node && Array.isArray(node.body)) {
		bodies.push(node.body);
	}
	for (const body of bodies) {
		for (const child of body) {
			if (isLeafStatement(child)) {
				const toks = statementTokens(source, child.span);
				if (toks.length > 0 && toks[0].kind === 'integerLiteral') {
					return true;
				}
				if (child.kind === 'Statement' && toks.length === 1
					&& (toks[0].kind === 'identifier' || toks[0].kind === 'keyword')
					&& source[child.span.end] === ':') {
					return true;
				}
			} else if (child.kind === 'ConditionalDirective') {
				return true;
			} else if (blockHasLandingPoint(source, child)) {
				return true;
			}
		}
	}
	return false;
}
