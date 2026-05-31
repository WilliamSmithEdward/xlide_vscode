// Member-access completion resolver.
//
// Given VBA source and a cursor offset positioned just after a member-access dot
// (e.g. `ThisWorkbook.`, `ws.Range("A1").`), this resolves the type of the
// receiver expression and returns the verified host members available on it.
//
// Pure analyzer code: depends only on the lexer, parser, and host model, never
// on vscode. See docs/xlide_vba_language_service_roadmap.md (Host-Context Member
// Completion addendum).

import { tokenize } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import { parseModule } from '../parser/parseModule';
import {
	BodyNode,
	ModuleNode,
	ProcedureNode,
	VariableGroupNode,
} from '../parser/nodes';
import type {
	HostMember,
	HostMemberKind,
	HostObjectModel,
} from '../host/excelObjectModel';
import {
	getHostMembers,
	resolveHostAlias,
	resolveHostGlobal,
	resolveMemberReturnType,
} from '../host/hostModel';
import { hasDocContent, renderDocMarkdown } from '../docs/docModel';
import type { VbaDoc } from '../docs/docModel';
import type { VbaProjectClassMembers } from '../symbols/symbolModel';

/** Project/module facts the resolver needs that come from outside the source. */
export interface MemberCompletionContext {
	/**
	 * Lowercased worksheet/document code name -> qualified host type, taken from
	 * the workbook's VBA project (e.g. "sheet1" -> "Excel.Worksheet",
	 * "thisworkbook" -> "Excel.Workbook"). Resolved by CODE NAME, not tab name.
	 */
	codeNames?: Record<string, string>;
	/** Qualified host type that `Me` resolves to in the current module. */
	meType?: string;
	/** Source-declared workbook class/UserForm/document members, keyed by type. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
}

/** A single member-completion result. */
export interface MemberCompletion {
	name: string;
	kind: HostMemberKind;
	/** Qualified type the member returns, when chainable. */
	returns?: string;
	/** True when source proves assignment to the member is allowed. */
	writable?: boolean;
	/** Declared value type accepted by assignment when source provides one. */
	writeType?: string;
	/** Qualified type the member belongs to (for detail text). */
	owner: string;
	/** Markdown documentation rendered from source-backed XML doc comments. */
	documentation?: string;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const PROJECT_TYPE_PREFIX = 'project:';

type CompletionMemberSource = Pick<HostMember, 'name' | 'kind' | 'returns'> & {
	writable?: boolean;
	writeType?: string;
	doc?: VbaDoc;
};

function word(token: VbaToken): string {
	return token.rawText;
}

function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

/** A logical-line boundary: a newline or a statement-separating colon. */
function isBoundary(token: VbaToken): boolean {
	return token.kind === 'newline' || token.rawText === ':';
}

/**
 * Resolves the member completions available at `offset`. Returns an empty array
 * when the cursor is not in a member-access position or the receiver type
 * cannot be resolved to a known host type.
 */
export function resolveMemberCompletions(
	source: string,
	offset: number,
	ctx: MemberCompletionContext = {},
): MemberCompletion[] {
	const prefixText = source.slice(0, Math.max(0, offset));
	// Keep newline tokens: they mark statement boundaries so a dangling
	// member-access dot on a previous line is not merged into this chain.
	const tokens = tokenize(prefixText).filter((t) => t.kind !== 'comment');
	if (tokens.length === 0) {
		return [];
	}

	// Identify the typed member prefix (text after the dot) and the dot itself.
	let i = tokens.length - 1;
	let typedPrefix = '';
	if (isIdentLike(tokens[i]) && i > 0 && tokens[i - 1].rawText === '.') {
		typedPrefix = tokens[i].rawText;
		i -= 1;
	}
	if (i < 0 || tokens[i].rawText !== '.') {
		return [];
	}
	// tokens[i] is the member-access dot; the receiver chain ends at i-1.
	const currentType = receiverTypeFromTokens(tokens, i, source, offset, ctx);
	if (!currentType) {
		return [];
	}

	const lowerPrefix = typedPrefix.toLowerCase();
	const surface = memberSurfaceForType(currentType, ctx);
	if (!surface) {
		return [];
	}
	return surface.members
		.filter((mem) => mem.name.toLowerCase().startsWith(lowerPrefix))
		.map((mem) => ({
			name: mem.name,
			kind: mem.kind,
			returns: mem.returns,
			writable: mem.writable,
			writeType: mem.writeType,
			owner: surface.owner,
			documentation: hasDocContent(mem.doc)
				? renderDocMarkdown(mem.doc)
				: undefined,
		}));
}

/**
 * Resolves the qualified host type whose members are accessible at a
 * member-access dot ending the text before `offset`. Returns undefined when the
 * cursor is not in a member-access position or the receiver cannot be resolved.
 *
 * Used by hover to describe `receiver.member` symbols; the dot may be followed
 * by a partially typed member name, which is ignored here.
 */
export function resolveReceiverTypeAt(
	source: string,
	offset: number,
	ctx: MemberCompletionContext = {},
): string | undefined {
	const prefixText = source.slice(0, Math.max(0, offset));
	const tokens = tokenize(prefixText).filter((t) => t.kind !== 'comment');
	if (tokens.length === 0) {
		return undefined;
	}
	let i = tokens.length - 1;
	if (isIdentLike(tokens[i]) && i > 0 && tokens[i - 1].rawText === '.') {
		i -= 1;
	}
	if (i < 0 || tokens[i].rawText !== '.') {
		return undefined;
	}
	return receiverTypeFromTokens(tokens, i, source, offset, ctx);
}

/**
 * Walks the receiver chain ending at the dot `tokens[dotIndex]` and resolves it
 * to a qualified host type, threading return types through each `.member` hop.
 */
function receiverTypeFromTokens(
	tokens: VbaToken[],
	dotIndex: number,
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	const chain = collectReceiverChain(tokens, dotIndex - 1);
	if (chain.length === 0) {
		return undefined;
	}
	const rootType = resolveRoot(chain[0], source, offset, ctx);
	if (!rootType) {
		return undefined;
	}
	let currentType: string | undefined = rootType;
	for (let s = 1; s < chain.length && currentType; s += 1) {
		currentType = resolveAnyMemberReturnType(currentType, chain[s], ctx);
	}
	return currentType;
}

/**
 * Walks left from `endIndex` collecting a dotted receiver chain of identifiers,
 * skipping balanced call/index parentheses. Returns the chain left-to-right,
 * e.g. ["ws", "Range", "Offset"] for `ws.Range("A1").Offset(1, 0)`. Returns an
 * empty array if the expression is not a simple member-access chain.
 */
function collectReceiverChain(tokens: VbaToken[], endIndex: number): string[] {
	const segments: string[] = [];
	let i = endIndex;
	for (;;) {
		// A statement boundary ends the receiver expression; anything to the left
		// belongs to a different statement and must not join this chain.
		if (i >= 0 && isBoundary(tokens[i])) {
			return [];
		}
		// Skip a trailing call/index argument list: ... ident ( args ) .
		if (i >= 0 && tokens[i].rawText === ')') {
			const open = matchParenLeft(tokens, i);
			if (open < 0) {
				return [];
			}
			i = open - 1;
			continue;
		}
		if (i < 0 || !isIdentLike(tokens[i])) {
			return [];
		}
		segments.unshift(word(tokens[i]));
		i -= 1;
		if (i >= 0 && tokens[i].rawText === '.') {
			i -= 1;
			continue;
		}
		break;
	}
	return segments;
}

/** Returns the index of the '(' matching the ')' at `closeIndex`, or -1. */
function matchParenLeft(tokens: VbaToken[], closeIndex: number): number {
	let depth = 0;
	for (let i = closeIndex; i >= 0; i -= 1) {
		const t = tokens[i].rawText;
		if (t === ')') {
			depth += 1;
		} else if (t === '(') {
			depth -= 1;
			if (depth === 0) {
				return i;
			}
		}
	}
	return -1;
}

/** Resolves the qualified host type of the root identifier of a chain. */
function resolveRoot(
	root: string,
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	const model = ctx.model;
	const lower = root.toLowerCase();

	if (lower === 'me') {
		return ctx.meType;
	}
	const asGlobal = resolveHostGlobal(root, model);
	if (asGlobal) {
		return asGlobal;
	}
	const asCode = ctx.codeNames?.[lower];
	if (asCode) {
		return asCode;
	}
	const declaredType = findDeclaredType(source, offset, root);
	if (declaredType) {
		return resolveDeclaredObjectType(declaredType, ctx, model);
	}
	return undefined;
}

function memberSurfaceForType(
	typeName: string,
	ctx: MemberCompletionContext,
): { owner: string; members: readonly CompletionMemberSource[] } | undefined {
	if (typeName.startsWith(PROJECT_TYPE_PREFIX)) {
		const projectType = projectClassMembersByName(ctx).get(
			typeName.slice(PROJECT_TYPE_PREFIX.length),
		);
		return projectType
			? { owner: projectType.name, members: projectType.members }
			: undefined;
	}
	return { owner: typeName, members: getHostMembers(typeName, ctx.model) };
}

function resolveAnyMemberReturnType(
	ownerType: string,
	memberName: string,
	ctx: MemberCompletionContext,
): string | undefined {
	if (!ownerType.startsWith(PROJECT_TYPE_PREFIX)) {
		return resolveMemberReturnType(ownerType, memberName, ctx.model);
	}
	const projectType = projectClassMembersByName(ctx).get(
		ownerType.slice(PROJECT_TYPE_PREFIX.length),
	);
	const member = projectType?.members.find(
		(m) => m.name.toLowerCase() === memberName.toLowerCase(),
	);
	return member?.returns
		? resolveDeclaredObjectType(member.returns, ctx, ctx.model)
		: undefined;
}

function resolveDeclaredObjectType(
	declaredType: string,
	ctx: MemberCompletionContext,
	model: HostObjectModel | undefined,
): string | undefined {
	const host = resolveHostAlias(declaredType, model);
	if (host) {
		return host;
	}
	const key = simpleTypeName(declaredType)?.toLowerCase();
	if (key && projectClassMembersByName(ctx).has(key)) {
		return projectTypeKey(key);
	}
	return undefined;
}

function simpleTypeName(typeText: string): string | undefined {
	const trimmed = typeText.trim();
	if (!IDENT_RE.test(trimmed)) {
		return undefined;
	}
	return trimmed;
}

function projectTypeKey(lowerName: string): string {
	return `${PROJECT_TYPE_PREFIX}${lowerName}`;
}

function projectClassMembersByName(
	ctx: MemberCompletionContext,
): ReadonlyMap<string, VbaProjectClassMembers> {
	const out = new Map<string, VbaProjectClassMembers>();
	const ambiguous = new Set<string>();
	for (const type of ctx.projectClassMembers ?? []) {
		const key = type.name.toLowerCase();
		if (ambiguous.has(key)) {
			continue;
		}
		if (out.has(key)) {
			out.delete(key);
			ambiguous.add(key);
			continue;
		}
		out.set(key, type);
	}
	return out;
}

/**
 * Finds the declared `As` type of a local variable, parameter, or module-level
 * variable named `name`, preferring the declaration in the procedure that
 * encloses `offset`. Returns the raw type text, or undefined.
 */
function findDeclaredType(
	source: string,
	offset: number,
	name: string,
): string | undefined {
	const module: ModuleNode = parseModule(source);
	const lower = name.toLowerCase();

	const enclosing = module.members.find(
		(mem): mem is ProcedureNode =>
			mem.kind === 'Procedure' &&
			offset >= mem.span.start &&
			offset <= mem.span.end,
	);

	if (enclosing) {
		for (const param of enclosing.params) {
			if (param.name.toLowerCase() === lower && param.asType) {
				return param.asType;
			}
		}
		const local = findInBody(enclosing.body, lower);
		if (local) {
			return local;
		}
	}

	for (const mem of module.members) {
		if (mem.kind === 'VariableGroup') {
			const hit = matchGroup(mem, lower);
			if (hit) {
				return hit;
			}
		}
	}
	return undefined;
}

/** Searches a procedure body (recursing into block nodes) for a declaration. */
function findInBody(body: BodyNode[], lower: string): string | undefined {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			const hit = matchGroup(node, lower);
			if (hit) {
				return hit;
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			const hit = findInBody(node.body, lower);
			if (hit) {
				return hit;
			}
		}
	}
	return undefined;
}

function matchGroup(group: VariableGroupNode, lower: string): string | undefined {
	for (const decl of group.declarations) {
		if (decl.name.toLowerCase() === lower && decl.asType) {
			return decl.asType;
		}
	}
	return undefined;
}
