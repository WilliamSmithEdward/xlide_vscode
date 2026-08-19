// Semantic tokens and lookup for type names.
//
// TextMate grammar can color only static tokens. This pure resolver marks parsed
// declaration type positions and `New` expressions after resolving the type
// name against project, VBA primitive, and host object-model type metadata.

import type {
	BodyNode,
	LeafStatementNode,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
	TypeFieldNode,
	VariableDeclNode,
	VariableGroupNode,
} from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenizeCached } from '../lexer/tokenize';
// statementTokensCached: per (source, span) LRU that derives statement
// tokens by slicing the shared cached module stream instead of re-lexing
// text slices. The collectors below visit every statement and re-visit each
// span across walkers, so the raw statementTokens re-lexed the whole module
// in pieces per pass - the type collector's dominant cost on large modules
// (measured 43 ms on the 947 KB corpus module, mostly lexing and the
// garbage it makes).
import { statementTokensCached as codeTokens, tokenName, tokenWord } from '../lexer/tokenHelpers';
import {
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostGlobalMember,
	resolveHostMember,
} from '../host/hostModel';
import type { HostObjectModel } from '../host/excelObjectModel';
import {
	resolveTypeName,
	type TypeCompletion,
	type TypeCompletionContext,
	type TypeCompletionKind,
} from '../completion/typeCompletion';
import { msFormsControlMembers } from '../completion/memberAccess';
import { VBA_USERFORM_TYPE as MSFORMS_USERFORM_TYPE } from '../host/userFormExtenderMembers';

export type TypeSemanticTokenType =
	| 'class'
	| 'enum'
	| 'enumMember'
	| 'struct'
	| 'type'
	| 'variable'
	| 'function';

export interface TypeSemanticToken {
	name: string;
	tokenType: TypeSemanticTokenType;
	span: Span;
	/** Standard semantic-token modifiers (e.g. `defaultLibrary` for host globals). */
	modifiers?: string[];
}

export interface ResolvedTypeReference extends TypeCompletion {
	span: Span;
	qualifier?: string;
	qualifierSpan?: Span;
	fullSpan?: Span;
}

export type TypeNameReferenceKind =
	| 'declaration'
	| 'newDeclaration'
	| 'newExpression'
	| 'typeOfIs'
	| 'implements';

export interface TypeNameReference {
	name: string;
	qualifier?: string;
	qualifierSpan?: Span;
	span: Span;
	fullSpan?: Span;
	kind: TypeNameReferenceKind;
}

function tokenTypeForCompletionKind(kind: TypeCompletionKind): TypeSemanticTokenType {
	switch (kind) {
		case 'class':
		case 'document':
		case 'userform':
		case 'host':
		case 'external':
			return 'class';
		case 'enum':
			return 'enum';
		case 'userType':
			return 'struct';
		case 'primitive':
		case 'ambiguous':
		case 'module':
			return 'type';
	}
}

function typeNameReferenceFromTokens(
	toks: readonly VbaToken[],
	typeIndex: number,
	base: number,
	kind: TypeNameReferenceKind,
): TypeNameReference | undefined {
	const firstName = tokenName(toks[typeIndex]);
	if (!firstName) {
		return undefined;
	}
	const firstSpan = {
		start: base + toks[typeIndex].start,
		end: base + toks[typeIndex].end,
	};
	const memberName = toks[typeIndex + 1]?.rawText === '.'
		? tokenName(toks[typeIndex + 2])
		: undefined;
	if (!memberName) {
		return {
			name: firstName,
			span: firstSpan,
			kind,
		};
	}
	const memberSpan = {
		start: base + toks[typeIndex + 2].start,
		end: base + toks[typeIndex + 2].end,
	};
	return {
		name: memberName,
		qualifier: firstName,
		qualifierSpan: firstSpan,
		span: memberSpan,
		fullSpan: {
			start: firstSpan.start,
			end: memberSpan.end,
		},
		kind,
	};
}

function typeNameSpanAfterAs(source: string, span: Span): TypeNameReference | undefined {
	const toks = codeTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'as') {
			continue;
		}
		let typeIndex = i + 1;
		let kind: TypeNameReferenceKind = 'declaration';
		if ((toks[typeIndex]?.canonicalText ?? toks[typeIndex]?.rawText ?? '').toLowerCase() === 'new') {
			typeIndex++;
			kind = 'newDeclaration';
		}
		const ref = typeNameReferenceFromTokens(toks, typeIndex, span.start, kind);
		if (ref) {
			return ref;
		}
	}
	return undefined;
}

function typeNameSpansAfterNew(source: string, span: Span): TypeNameReference[] {
	const toks = codeTokens(source, span);
	const out: TypeNameReference[] = [];
	for (let i = 0; i < toks.length; i++) {
		if ((toks[i].canonicalText ?? toks[i].rawText).toLowerCase() !== 'new') {
			continue;
		}
		const ref = typeNameReferenceFromTokens(toks, i + 1, span.start, 'newExpression');
		if (ref) {
			out.push(ref);
		}
	}
	return out;
}

function typeNameSpansAfterTypeOfIs(source: string, span: Span): TypeNameReference[] {
	const toks = codeTokens(source, span);
	const out: TypeNameReference[] = [];
	let sawTypeOf = false;
	for (let i = 0; i < toks.length; i++) {
		const lower = (toks[i].canonicalText ?? toks[i].rawText).toLowerCase();
		if (lower === 'typeof') {
			sawTypeOf = true;
			continue;
		}
		if (!sawTypeOf) {
			continue;
		}
		if (lower === 'is') {
			const ref = typeNameReferenceFromTokens(toks, i + 1, span.start, 'typeOfIs');
			if (ref) {
				out.push(ref);
			}
			sawTypeOf = false;
			continue;
		}
		// `TypeOf expr Is Type` only allows the operand's qualified chain
		// (name, '.', name, ...) between TypeOf and Is. Any other significant
		// token means this is a malformed/non-TypeOf use; clear the flag so a
		// later unrelated `Is` in the same statement is not misclassified.
		if (lower !== '.' && !tokenName(toks[i])) {
			sawTypeOf = false;
		}
	}
	return out;
}

function headerSpanForProcedure(source: string, proc: ProcedureNode): Span {
	const nl = source.indexOf('\n', proc.span.start);
	return {
		start: proc.span.start,
		end: nl === -1 ? proc.span.end : Math.min(nl, proc.span.end),
	};
}

function returnTypeNameSpan(source: string, proc: ProcedureNode): TypeNameReference | undefined {
	const header = headerSpanForProcedure(source, proc);
	const toks = codeTokens(source, header);
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
			continue;
		}
		if (raw === ')') {
			depth--;
			continue;
		}
		if (depth !== 0 || (toks[i].canonicalText ?? raw).toLowerCase() !== 'as') {
			continue;
		}
		const ref = typeNameReferenceFromTokens(toks, i + 1, header.start, 'declaration');
		if (!ref) {
			return undefined;
		}
		return ref;
	}
	return undefined;
}

function collectVariableGroup(
	source: string,
	group: VariableGroupNode,
	out: TypeNameReference[],
): void {
	for (const decl of group.declarations) {
		collectVariableDecl(source, decl, out);
	}
}

function collectVariableDecl(
	source: string,
	decl: VariableDeclNode,
	out: TypeNameReference[],
): void {
	if (!decl.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, decl.span));
}

function collectTypeField(
	source: string,
	field: TypeFieldNode,
	out: TypeNameReference[],
): void {
	if (!field.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, field.span));
}

function collectParameter(
	source: string,
	param: ParameterNode,
	out: TypeNameReference[],
): void {
	if (!param.asType) {
		return;
	}
	pushTypeHit(out, typeNameSpanAfterAs(source, param.span));
}

function collectBody(
	source: string,
	body: BodyNode[],
	out: TypeNameReference[],
): void {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			collectVariableGroup(source, node, out);
		} else if (isLeafStatement(node)) {
			collectStatement(source, node, out);
		} else if ('body' in node && Array.isArray(node.body)) {
			// A block holds its opening header (an If/Do While/While/For/Select
			// condition) outside `body`, so scan the header span for TypeOf...Is /
			// New type references before recursing into the body statements.
			collectBlockHeader(source, node, out);
			collectBody(source, node.body, out);
		}
	}
}

function collectStatement(
	source: string,
	stmt: LeafStatementNode,
	out: TypeNameReference[],
): void {
	collectHeaderSpan(source, stmt.span, out);
}

function collectBlockHeader(
	source: string,
	node: BodyNode,
	out: TypeNameReference[],
): void {
	// The leading `If` condition lives in branches[0].headerSpan; ElseIf/Else
	// headers are already walked as raw body statements, so only the opener is
	// missed. Other blocks keep their opener as the block's first physical line.
	if (node.kind === 'IfBlock') {
		const header = node.branches[0]?.headerSpan;
		if (header) {
			collectHeaderSpan(source, header, out);
		}
		return;
	}
	const block = node as { span?: Span; body?: { span: Span }[] };
	if (block.span) {
		// Bound the header by the first body statement rather than the first physical
		// newline, so a line-continued header (`For Each x In _` / `New T`) is fully
		// covered and the scan never overlaps the body (which collectBody walks),
		// avoiding duplicate tokens.
		const end = block.body && block.body.length > 0
			? block.body[0].span.start
			: block.span.end;
		collectHeaderSpan(source, { start: block.span.start, end }, out);
	}
}

function collectHeaderSpan(source: string, span: Span, out: TypeNameReference[]): void {
	for (const hit of typeNameSpansAfterNew(source, span)) {
		out.push(hit);
	}
	for (const hit of typeNameSpansAfterTypeOfIs(source, span)) {
		out.push(hit);
	}
}

function collectImplements(source: string, out: TypeNameReference[], scanEnd: number = source.length): void {
	let lineStart = 0;
	while (lineStart <= scanEnd) {
		let lineEnd = source.indexOf('\n', lineStart);
		if (lineEnd < 0) {
			lineEnd = source.length;
		}
		let line = source.slice(lineStart, lineEnd);
		if (line.endsWith('\r')) {
			line = line.slice(0, -1);
		}
		const code = line.replace(/'.*$/, '');
		const match = /^\s*Implements\s+([\p{L}_][\p{L}\p{M}\p{N}_]*(?:\.[\p{L}_][\p{L}\p{M}\p{N}_]*)?)/iu.exec(code);
		if (match) {
			const rawName = match[1];
			const column = line.indexOf(rawName, match.index);
			const dot = rawName.indexOf('.');
			if (column >= 0 && dot > 0) {
				out.push({
					name: rawName.slice(dot + 1),
					qualifier: rawName.slice(0, dot),
					qualifierSpan: {
						start: lineStart + column,
						end: lineStart + column + dot,
					},
					span: {
						start: lineStart + column + dot + 1,
						end: lineStart + column + rawName.length,
					},
					fullSpan: {
						start: lineStart + column,
						end: lineStart + column + rawName.length,
					},
					kind: 'implements',
				});
			} else if (column >= 0) {
				out.push({
					name: rawName,
					span: {
						start: lineStart + column,
						end: lineStart + column + rawName.length,
					},
					kind: 'implements',
				});
			}
		}
		if (lineEnd === source.length) {
			break;
		}
		lineStart = lineEnd + 1;
	}
}

function collectProcedure(
	source: string,
	proc: ProcedureNode,
	out: TypeNameReference[],
): void {
	for (const param of proc.params) {
		collectParameter(source, param, out);
	}
	if (proc.returnType) {
		pushTypeHit(out, returnTypeNameSpan(source, proc));
	}
	collectBody(source, proc.body, out);
}

function collectModule(
	source: string,
	mod: ModuleNode,
): TypeNameReference[] {
	const out: TypeNameReference[] = [];
	// `Implements` is only legal in the declarations section, before any
	// procedure, so the line scan stops at the first procedure instead of
	// walking every line of the module body.
	const firstProcedureStart = mod.members.find((m) => m.kind === 'Procedure')?.span.start
		?? source.length;
	collectImplements(source, out, firstProcedureStart);
	for (const member of mod.members) {
		if (member.kind === 'VariableGroup') {
			collectVariableGroup(source, member, out);
		} else if (member.kind === 'Type') {
			for (const field of member.fields) {
				collectTypeField(source, field, out);
			}
		} else if (member.kind === 'Procedure') {
			collectProcedure(source, member, out);
		}
	}
	return out.sort((a, b) => a.span.start - b.span.start || a.span.end - b.span.end);
}

function pushTypeHit(
	out: TypeNameReference[],
	hit: TypeNameReference | undefined,
): void {
	if (hit) {
		out.push(hit);
	}
}

function semanticTokenForHit(
	ctx: TypeCompletionContext,
	hit: TypeNameReference,
): TypeSemanticToken | undefined {
	const resolved = resolveTypeName(typeReferenceLookupName(hit), ctx);
	if (!resolved) {
		return undefined;
	}
	return {
		name: hit.name,
		tokenType: tokenTypeForCompletionKind(resolved.kind),
		span: hit.span,
	};
}

export function typeReferenceLookupName(hit: TypeNameReference): string {
	return hit.qualifier ? `${hit.qualifier}.${hit.name}` : hit.name;
}

export function collectTypeNameReferences(source: string): TypeNameReference[] {
	return collectModule(source, parseModule(source));
}

// Token words that put the following identifier in a TYPE position, where a
// host name like `Application` is the type (colored by the type collector as
// `class`) rather than the host-global value. Excluded from value coloring.
const TYPE_POSITION_LEADS = new Set(['as', 'new']);

/**
 * Names declared anywhere in the module (procedures, parameters, variables,
 * constants, types, enum members, Declares). A host-global name that is also
 * declared here is shadowed, so it must not be colored as the host global.
 */
function collectModuleDeclaredNames(module: ModuleNode): Set<string> {
	const names = new Set<string>();
	const addBody = (body: BodyNode[]): void => {
		for (const node of body) {
			if (node.kind === 'VariableGroup') {
				for (const decl of node.declarations) {
					names.add(decl.name.toLowerCase());
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				addBody(node.body);
			}
		}
	};
	for (const member of module.members) {
		switch (member.kind) {
			case 'Procedure':
				names.add(member.name.toLowerCase());
				for (const param of member.params) {
					names.add(param.name.toLowerCase());
				}
				addBody(member.body);
				break;
			case 'VariableGroup':
				for (const decl of member.declarations) {
					names.add(decl.name.toLowerCase());
				}
				break;
			case 'Type':
			case 'Enum':
			case 'Declare':
			case 'Event':
				names.add(member.name.toLowerCase());
				if (member.kind === 'Enum') {
					for (const enumMember of member.members) {
						names.add(enumMember.name.toLowerCase());
					}
				}
				break;
			default:
				break;
		}
	}
	return names;
}

/**
 * Bare host names in VALUE position: injected globals (`Application`,
 * `ActiveSheet`) as `variable.defaultLibrary`, the hidden Global interface's
 * members (issue #34: a method paints as `function`, a property as a host
 * value), and resolved host enum constants as `enumMember` (issue #35: xlUp
 * and xlLandscape in one tier). A name is colored only when it (1) resolves
 * in the module's own model, (2) is not a member access (`x.Application`),
 * (3) is not in a type position (`As Application` / `New Application`,
 * already colored as a type), and (4) is not shadowed by an in-module
 * declaration or a designer-declared control (issue #30: inside a form the
 * control wins name binding). The conservative gating keeps a local named
 * the same as a host name from being mis-colored.
 */
export function collectHostGlobalTokens(
	source: string,
	model?: HostObjectModel,
	implicitMembers?: readonly { name: string }[],
): TypeSemanticToken[] {
	const declared = collectModuleDeclaredNames(parseModule(source));
	for (const member of implicitMembers ?? []) {
		declared.add(member.name.toLowerCase());
	}
	// tokenizeCached: this runs in one semantic-token pass with the other
	// collectors over the same source string, and a raw tokenize of a large
	// module costs tens of milliseconds per duplicate (measured 36 ms on the
	// 947 KB corpus module). The filter builds a fresh array, so the shared
	// cached token stream is never mutated.
	const tokens = tokenizeCached(source).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const out: TypeSemanticToken[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		if (tok.kind !== 'identifier') {
			continue;
		}
		const lower = tok.rawText.toLowerCase();
		if (declared.has(lower)) {
			continue;
		}
		// Resolved against the module's own host model (issue #24): Word's
		// ActiveDocument paints in a Word module, and Excel's ActiveSheet
		// does not. A name that is not an injected global may still be a
		// member of the host's hidden Global interface, callable bare -
		// Word's InchesToPoints, Excel's Union (issue #34) - or a host enum
		// constant, which paints as the enum member it is (issue #35: xlUp
		// and xlLandscape read as one tier, not the grammar's curated few).
		const globalType = resolveHostGlobal(tok.rawText, model);
		const globalMember = globalType ? undefined : resolveHostGlobalMember(tok.rawText, model);
		const constant = globalType || globalMember
			? undefined
			: resolveHostConstant(tok.rawText, model);
		if (!globalType && !globalMember && !constant) {
			continue;
		}
		const prev = tokens[i - 1];
		if (prev) {
			if (prev.rawText === '.' || prev.rawText === '!') {
				continue; // member / bang access on another receiver
			}
			if (prev.kind === 'keyword' && TYPE_POSITION_LEADS.has(tokenWord(prev))) {
				continue; // type position, owned by the type collector
			}
		}
		const span = { start: tok.start, end: tok.end };
		if (constant) {
			out.push({ name: tok.rawText, tokenType: 'enumMember', span });
		} else if (globalMember?.kind === 'method') {
			// A bare Global method is a call, painted the way member method
			// calls are (#20/#29); Global properties are host-injected
			// values, tinted the way the injected globals are.
			out.push({ name: tok.rawText, tokenType: 'function', span });
		} else {
			out.push({
				name: tok.rawText,
				tokenType: 'variable',
				span,
				modifiers: ['defaultLibrary'],
			});
		}
	}
	return out;
}

/** What the implicit-member method collector needs from outside the source. */
export interface ImplicitMemberTokenContext {
	/** The form's designer-declared controls: name plus MSForms type. */
	implicitMembers?: readonly { name: string; type: string }[];
	/**
	 * What `Me` denotes. Optional: a module with designer-declared controls is
	 * a form, so `Me` is taken as `MSForms.UserForm` when this is omitted and
	 * controls are present. Pass it to say otherwise - any non-MSForms type
	 * (a document module's `Excel.Workbook`) leaves `Me.` alone entirely.
	 */
	meType?: string;
}

/**
 * Method calls on a form's implicit members, for `function` semantic coloring:
 * `RegionPick.AddItem` paints the way `Len` does, while `Taxable.Value` (a
 * property) and `RegionPick.NotAMember` (unresolved) stay untouched (issue
 * #20). Resolved only, never guessed: the receiver must be a designer-declared
 * control (or `Me` in a form), not shadowed by any in-module declaration, and
 * the member must be a method the MSForms surface actually carries.
 *
 * Receivers themselves are deliberately not colored here.
 */
export function collectImplicitMemberMethodTokens(
	source: string,
	ctx: ImplicitMemberTokenContext = {},
): TypeSemanticToken[] {
	const controls = new Map(
		(ctx.implicitMembers ?? []).map((member) => [member.name.toLowerCase(), member.type]),
	);
	// Controls exist only on a form, so their presence answers what `Me` is
	// without the caller having to say.
	const declaredMeType = ctx.meType ?? (controls.size > 0 ? MSFORMS_USERFORM_TYPE : undefined);
	const meType = declaredMeType && declaredMeType.startsWith('MSForms.')
		? declaredMeType
		: undefined;
	if (controls.size === 0 && !meType) {
		return [];
	}
	const declared = collectModuleDeclaredNames(parseModule(source));
	// Newlines are KEPT. A dot that opens a line is a `With` block's member
	// access, not a member of whatever the line above happened to end with, so
	// dropping newlines here would paint `.Clear` after a line ending in a
	// control name - a guess, and the one thing this must never do.
	// (tokenizeCached for the same reason as collectHostGlobalTokens; the
	// filter keeps the shared stream unmutated.)
	const tokens = tokenizeCached(source).filter((t) => t.kind !== 'comment');
	const out: TypeSemanticToken[] = [];
	for (let i = 2; i < tokens.length; i++) {
		const tok = tokens[i];
		if (tok.kind !== 'identifier' || tokens[i - 1].rawText !== '.') {
			continue;
		}
		const ownerType = implicitReceiverType(tokens, i - 2, controls, declared, meType);
		if (!ownerType) {
			continue;
		}
		const member = msFormsControlMembers(ownerType)?.find(
			(candidate) => candidate.name.toLowerCase() === tok.rawText.toLowerCase(),
		);
		if (member?.kind !== 'method') {
			continue;
		}
		out.push({
			name: tok.rawText,
			tokenType: 'function',
			span: { start: tok.start, end: tok.end },
		});
	}
	return out;
}

/**
 * The MSForms type of the receiver ending at `recvIndex`, when that receiver
 * is a form's implicit member (`RegionPick.`, `Me.RegionPick.`) or `Me` in a
 * form - undefined for every other shape, including a receiver shadowed by an
 * in-module declaration and any longer chain.
 */
function implicitReceiverType(
	tokens: readonly VbaToken[],
	recvIndex: number,
	controls: ReadonlyMap<string, string>,
	declared: ReadonlySet<string>,
	meType: string | undefined,
): string | undefined {
	const recv = tokens[recvIndex];
	if (!recv) {
		return undefined;
	}
	if (tokenWord(recv) === 'me') {
		return meType;
	}
	if (recv.kind !== 'identifier') {
		return undefined;
	}
	const name = recv.rawText.toLowerCase();
	const before = tokens[recvIndex - 1];
	if (!before || (before.rawText !== '.' && before.rawText !== '!')) {
		// Chain root: the bare control name, unless a declaration shadows it.
		return declared.has(name) ? undefined : controls.get(name);
	}
	// `Me.RegionPick.Member` is the same control; any other qualifier is not.
	if (
		before.rawText === '.' &&
		tokenWord(tokens[recvIndex - 2]) === 'me' &&
		meType !== undefined
	) {
		return controls.get(name);
	}
	return undefined;
}

/** What the host member method collector needs from outside the source. */
export interface HostMemberTokenContext {
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
	/**
	 * Lowercased document code name -> qualified host type ("sheet1" ->
	 * "Excel.Worksheet", "thisdocument" -> "Word.Document"), so a code-name
	 * receiver resolves the way a host global does.
	 */
	codeNames?: Record<string, string>;
	/**
	 * A form's designer-declared controls. Inside the form their names bind
	 * before the host globals, so a control named like a global shadows it.
	 */
	implicitMembers?: readonly { name: string }[];
	/**
	 * Qualified host type `Me` denotes in this module - a document module's
	 * "Excel.Worksheet" / "Excel.Workbook" / "Word.Document" - so `Me.Calculate`
	 * paints the way `Sheet1.Calculate` does (issue #31). Absent (or a
	 * non-host type such as MSForms.UserForm, whose members the implicit
	 * collector owns) leaves `Me.` alone.
	 */
	meType?: string;
	/**
	 * Project type names visible to the module. A project class named like a
	 * host type wins the `As` clause, so a local declared with that name must
	 * not resolve as the host type (issue #33).
	 */
	projectTypes?: readonly { name: string }[];
}

/**
 * Method calls on a host receiver, for `function` semantic coloring:
 * `ActiveSheet.Calculate` and Word's `ActiveDocument.FitToPages` paint the
 * way `RegionPick.AddItem` does (issue #29, extending issue #20's
 * convention). Resolved only, never guessed: the receiver must be a host
 * global or a document code name at the chain root, shadowed by nothing in
 * the module, and the member must resolve to a method on the receiver's host
 * type. Properties and unresolved members take no token, and longer chains
 * stay out of scope the way they do for controls.
 */
export function collectHostMemberMethodTokens(
	source: string,
	ctx: HostMemberTokenContext = {},
): TypeSemanticToken[] {
	const declared = collectDeclaredNameHostTypes(parseModule(source), ctx);
	for (const member of ctx.implicitMembers ?? []) {
		// Controls shadow the host reading; their members are the implicit
		// collector's to paint.
		declared.set(member.name.toLowerCase(), null);
	}
	const codeNames = ctx.codeNames ?? {};
	// Newlines are KEPT for the same reason as collectImplicitMemberMethodTokens:
	// a dot that opens a line is a With block's member access, not a member of
	// whatever the line above happened to end with. (tokenizeCached; the filter
	// keeps the shared stream unmutated.)
	const tokens = tokenizeCached(source).filter((t) => t.kind !== 'comment');
	const out: TypeSemanticToken[] = [];
	for (let i = 2; i < tokens.length; i++) {
		const tok = tokens[i];
		if (tok.kind !== 'identifier' || tokens[i - 1].rawText !== '.') {
			continue;
		}
		const recv = tokens[i - 2];
		const before = tokens[i - 3];
		if (before && (before.rawText === '.' || before.rawText === '!')) {
			continue; // not the chain root
		}
		const receiverType = hostReceiverType(recv, ctx, declared, codeNames);
		if (!receiverType) {
			continue;
		}
		const member = resolveHostMember(receiverType, tok.rawText, ctx.model);
		if (member?.kind !== 'method') {
			continue;
		}
		out.push({
			name: tok.rawText,
			tokenType: 'function',
			span: { start: tok.start, end: tok.end },
		});
	}
	return out;
}

/**
 * The qualified host type of a chain-root receiver token: `Me` when the
 * module's `Me` denotes a host document type; a declared local whose `As`
 * clause resolves to a host object type (issue #33); otherwise an unshadowed
 * host global, Global-interface member, or document code name. Undefined for
 * every other shape.
 */
function hostReceiverType(
	recv: VbaToken | undefined,
	ctx: HostMemberTokenContext,
	declared: ReadonlyMap<string, string | null>,
	codeNames: Readonly<Record<string, string>>,
): string | undefined {
	if (!recv) {
		return undefined;
	}
	if (tokenWord(recv) === 'me') {
		// `Me` cannot be shadowed; a non-host meType simply resolves no member.
		return ctx.meType;
	}
	if (recv.kind !== 'identifier') {
		return undefined;
	}
	const lower = recv.rawText.toLowerCase();
	const declaredType = declared.get(lower);
	if (declaredType !== undefined) {
		// A declared name binds first: typed by its host `As` clause, or a
		// shadow that ends the host reading (null).
		return declaredType ?? undefined;
	}
	return resolveHostGlobal(recv.rawText, ctx.model)
		?? resolveHostGlobalMember(recv.rawText, ctx.model)?.returns
		?? codeNames[lower];
}

/**
 * Every declared name in the module, mapped to the qualified host object type
 * its `As` clause resolves to - or null when it declares no host type at all
 * (untyped, primitive, project-shadowed, a procedure or enum name) or when
 * two declarations of the name disagree. The map is deliberately whole-module
 * rather than scope-aware: a name typed the same everywhere paints, a name
 * that means different things in different procedures stays plain (issue #33).
 */
function collectDeclaredNameHostTypes(
	module: ModuleNode,
	ctx: HostMemberTokenContext,
): Map<string, string | null> {
	const projectNames = new Set(
		(ctx.projectTypes ?? []).map((type) => type.name.toLowerCase()),
	);
	const out = new Map<string, string | null>();
	const merge = (name: string, asType: string | undefined): void => {
		const resolved = asType && !projectNames.has(asType.trim().toLowerCase())
			? resolveHostAlias(asType, ctx.model) ?? null
			: null;
		const existing = out.get(name.toLowerCase());
		if (existing === undefined) {
			out.set(name.toLowerCase(), resolved);
		} else if (existing !== resolved) {
			out.set(name.toLowerCase(), null);
		}
	};
	const addBody = (body: BodyNode[]): void => {
		for (const node of body) {
			if (node.kind === 'VariableGroup') {
				for (const decl of node.declarations) {
					merge(decl.name, decl.asType);
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				addBody(node.body);
			}
		}
	};
	for (const member of module.members) {
		switch (member.kind) {
			case 'Procedure':
				merge(member.name, undefined);
				for (const param of member.params) {
					merge(param.name, param.asType);
				}
				addBody(member.body);
				break;
			case 'VariableGroup':
				for (const decl of member.declarations) {
					merge(decl.name, decl.asType);
				}
				break;
			case 'Type':
			case 'Enum':
			case 'Declare':
			case 'Event':
				merge(member.name, undefined);
				if (member.kind === 'Enum') {
					for (const enumMember of member.members) {
						merge(enumMember.name, undefined);
					}
				}
				break;
			default:
				break;
		}
	}
	return out;
}

export function resolveTypeSemanticTokens(
	source: string,
	ctx: TypeCompletionContext = {},
): TypeSemanticToken[] {
	return collectTypeNameReferences(source)
		.map((hit) => semanticTokenForHit(ctx, hit))
		.filter((token): token is TypeSemanticToken => Boolean(token));
}

export function resolveTypeReferenceAt(
	source: string,
	offset: number,
	ctx: TypeCompletionContext = {},
): ResolvedTypeReference | undefined {
	const hit = collectTypeNameReferences(source).find(
		(candidate) => offset >= candidate.span.start && offset <= candidate.span.end,
	);
	if (!hit) {
		return undefined;
	}
	const resolved = resolveTypeName(typeReferenceLookupName(hit), ctx);
	if (!resolved) {
		return undefined;
	}
	return {
		...resolved,
		span: hit.span,
		qualifier: hit.qualifier,
		qualifierSpan: hit.qualifierSpan,
		fullSpan: hit.fullSpan,
	};
}
