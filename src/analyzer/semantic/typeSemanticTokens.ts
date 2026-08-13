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
import { tokenize } from '../lexer/tokenize';
import { statementTokens as codeTokens, tokenName, tokenWord } from '../lexer/tokenHelpers';
import { resolveHostGlobal } from '../host/hostModel';
import {
	resolveTypeName,
	type TypeCompletion,
	type TypeCompletionContext,
	type TypeCompletionKind,
} from '../completion/typeCompletion';
import { msFormsControlMembers } from '../completion/memberAccess';

export type TypeSemanticTokenType = 'class' | 'enum' | 'struct' | 'type' | 'variable' | 'function';

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

function collectImplements(source: string, out: TypeNameReference[]): void {
	let lineStart = 0;
	while (lineStart <= source.length) {
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
	collectImplements(source, out);
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
 * Host-injected globals (`Application`, `ActiveSheet`, `ThisWorkbook`, ...) used
 * in VALUE position, for `variable.defaultLibrary` semantic coloring. A name is
 * colored only when it (1) resolves to a known host global, (2) is not a member
 * access (`x.Application`), (3) is not in a type position (`As Application` /
 * `New Application`, already colored as a type), and (4) is not shadowed by an
 * in-module declaration. The conservative gating keeps a local named the same
 * as a host global from being mis-colored.
 */
export function collectHostGlobalTokens(source: string): TypeSemanticToken[] {
	const declared = collectModuleDeclaredNames(parseModule(source));
	const tokens = tokenize(source).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const out: TypeSemanticToken[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const tok = tokens[i];
		if (tok.kind !== 'identifier') {
			continue;
		}
		const lower = tok.rawText.toLowerCase();
		if (declared.has(lower) || !resolveHostGlobal(tok.rawText)) {
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
		out.push({
			name: tok.rawText,
			tokenType: 'variable',
			span: { start: tok.start, end: tok.end },
			modifiers: ['defaultLibrary'],
		});
	}
	return out;
}

/** What the implicit-member method collector needs from outside the source. */
export interface ImplicitMemberTokenContext {
	/** The form's designer-declared controls: name plus MSForms type. */
	implicitMembers?: readonly { name: string; type: string }[];
	/** What `Me` denotes; only an `MSForms.*` type engages the collector for `Me`. */
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
	const meType = ctx.meType && /^MSForms\./.test(ctx.meType) ? ctx.meType : undefined;
	if (controls.size === 0 && !meType) {
		return [];
	}
	const declared = collectModuleDeclaredNames(parseModule(source));
	const tokens = tokenize(source).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
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
