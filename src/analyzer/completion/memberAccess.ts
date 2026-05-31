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
	StatementNode,
	VariableGroupNode,
} from '../parser/nodes';
import type {
	HostMember,
	HostMemberKind,
	HostObjectModel,
} from '../host/excelObjectModel';
import {
	getHostMembers,
	getHostType,
	resolveHostMemberSignature,
	resolveHostAlias,
	resolveHostGlobal,
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
	/** Verified call signature, when the host metadata has one. */
	signature?: string;
	/** True when source proves assignment to the member is allowed. */
	writable?: boolean;
	/** Declared value type accepted by assignment when source provides one. */
	writeType?: string;
	/** Qualified type the member belongs to (for detail text). */
	owner: string;
	/** True when the owner member surface is complete enough to prove absence. */
	surfaceExhaustive?: boolean;
	/** Markdown documentation rendered from source-backed XML doc comments. */
	documentation?: string;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const PROJECT_TYPE_PREFIX = 'project:';
const COMBINED_TYPE_PREFIX = 'combined:';
const COMBINED_TYPE_SEPARATOR = '|';
const UNION_TYPE_PREFIX = 'union:';
const UNION_TYPE_SEPARATOR = '|';

type CompletionMemberSource = Pick<HostMember, 'name' | 'kind' | 'returns'> & {
	writable?: boolean;
	writeType?: string;
	doc?: VbaDoc;
};

interface MemberSurface {
	owner: string;
	members: readonly CompletionMemberSource[];
	exhaustive: boolean;
}

interface ReceiverChainSegment {
	name: string;
	hasArguments: boolean;
}

interface ReceiverChain {
	segments: ReceiverChainSegment[];
	startIndex: number;
}

interface ResolvedMemberReturn {
	type: string;
	kind: HostMemberKind;
}

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
			signature: signatureForMember(currentType, mem.name, ctx),
			writable: mem.writable,
			writeType: mem.writeType,
			owner: surface.owner,
			surfaceExhaustive: surface.exhaustive,
			documentation: hasDocContent(mem.doc)
				? renderDocMarkdown(mem.doc)
				: undefined,
		}));
}

function signatureForMember(
	typeName: string,
	memberName: string,
	ctx: MemberCompletionContext,
): string | undefined {
	const union = parseUnionTypeKey(typeName);
	if (union) {
		const signatures = union
			.map((type) => signatureForMember(type, memberName, ctx))
			.filter((signature): signature is string => Boolean(signature));
		const distinct = new Set(signatures);
		return distinct.size === 1 ? signatures[0] : undefined;
	}
	const combined = parseCombinedTypeKey(typeName);
	if (combined) {
		return resolveHostMemberSignature(combined.hostType, memberName, ctx.model);
	}
	if (typeName.startsWith(PROJECT_TYPE_PREFIX)) {
		return undefined;
	}
	return resolveHostMemberSignature(typeName, memberName, ctx.model);
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
	return receiverTypeFromChain(chain, source, offset, ctx);
}

function receiverTypeFromExpressionTokens(
	tokens: VbaToken[],
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	if (tokens.length === 0) {
		return undefined;
	}
	const chain = collectReceiverChainWithStart(tokens, tokens.length - 1);
	if (!chain) {
		return undefined;
	}
	const prefix = tokens.slice(0, chain.startIndex);
	if (
		prefix.length > 0 &&
		!(prefix.length === 1 && prefix[0].rawText.toLowerCase() === 'new')
	) {
		return undefined;
	}
	return receiverTypeFromChain(chain.segments, source, offset, ctx);
}

function receiverTypeFromChain(
	chain: ReceiverChainSegment[],
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	if (chain.length === 0) {
		return undefined;
	}
	const root = chain[0];
	const rootType = resolveRoot(root.name, source, offset, ctx);
	if (!rootType) {
		return undefined;
	}
	let currentType = applyDefaultMemberReturnType(rootType, root.hasArguments, ctx);
	for (let s = 1; s < chain.length && currentType; s += 1) {
		const segment = chain[s];
		const resolved = resolveAnyMemberReturnType(currentType, segment.name, ctx);
		if (!resolved) {
			return undefined;
		}
		currentType = applyDefaultMemberReturnType(
			resolved.type,
			segment.hasArguments && resolved.kind !== 'method',
			ctx,
		);
	}
	return currentType;
}

/**
 * Walks left from `endIndex` collecting a dotted receiver chain of identifiers,
 * skipping balanced call/index parentheses. Returns segments left-to-right,
 * e.g. ws, Range(args), Offset(args) for `ws.Range("A1").Offset(1, 0)`. Returns
 * an empty array if the expression is not a simple member-access chain.
 */
function collectReceiverChain(
	tokens: VbaToken[],
	endIndex: number,
): ReceiverChainSegment[] {
	return collectReceiverChainWithStart(tokens, endIndex)?.segments ?? [];
}

function collectReceiverChainWithStart(
	tokens: VbaToken[],
	endIndex: number,
): ReceiverChain | undefined {
	const segments: ReceiverChainSegment[] = [];
	let i = endIndex;
	let pendingHasArguments = false;
	let startIndex = -1;
	for (;;) {
		// A statement boundary ends the receiver expression; anything to the left
		// belongs to a different statement and must not join this chain.
		if (i >= 0 && isBoundary(tokens[i])) {
			return undefined;
		}
		// Skip a trailing call/index argument list: ... ident ( args ) .
		if (i >= 0 && tokens[i].rawText === ')') {
			const open = matchParenLeft(tokens, i);
			if (open < 0) {
				return undefined;
			}
			pendingHasArguments = true;
			i = open - 1;
			continue;
		}
		if (i < 0 || !isIdentLike(tokens[i])) {
			return undefined;
		}
		startIndex = i;
		segments.unshift({
			name: word(tokens[i]),
			hasArguments: pendingHasArguments,
		});
		pendingHasArguments = false;
		i -= 1;
		if (i >= 0 && tokens[i].rawText === '.') {
			i -= 1;
			continue;
		}
		break;
	}
	return { segments, startIndex };
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
	const projectKey = projectClassMembersByName(ctx).has(lower) ? lower : undefined;
	const asGlobal = resolveHostGlobal(root, model);
	if (asGlobal) {
		return projectKey ? combinedTypeKey(projectKey, asGlobal) : asGlobal;
	}
	const asCode = ctx.codeNames?.[lower];
	if (asCode) {
		return projectKey ? combinedTypeKey(projectKey, asCode) : asCode;
	}
	const declaredType = findDeclaredType(source, offset, root);
	if (declaredType) {
		const declaredObjectType = resolveDeclaredObjectType(declaredType, ctx, model);
		if (declaredObjectType) {
			return declaredObjectType;
		}
		if (!isGenericObjectDeclaration(declaredType)) {
			return undefined;
		}
	}
	return findSetAssignedObjectType(source, offset, root, ctx);
}

function memberSurfaceForType(
	typeName: string,
	ctx: MemberCompletionContext,
): MemberSurface | undefined {
	const union = parseUnionTypeKey(typeName);
	if (union) {
		const surfaces = union
			.map((item) => memberSurfaceForType(item, ctx))
			.filter((item): item is MemberSurface => Boolean(item));
		if (surfaces.length === 0) {
			return undefined;
		}
		return {
			owner: union.map(displayTypeName).join(' | '),
			members: mergeCompletionMembers(...surfaces.map((surface) => surface.members)),
			exhaustive: surfaces.every((surface) => surface.exhaustive),
		};
	}
	const combined = parseCombinedTypeKey(typeName);
	if (combined) {
		const projectType = projectClassMembersByName(ctx).get(combined.projectKey);
		const hostType = getHostType(combined.hostType, ctx.model);
		if (!projectType && !hostType) {
			return undefined;
		}
		return {
			owner: projectType?.name ?? combined.hostType,
			members: mergeCompletionMembers(
				projectType?.members ?? [],
				getHostMembers(combined.hostType, ctx.model),
			),
			exhaustive:
				projectSourceSurfaceCompleteWhenMergedWithHost(projectType) &&
				hostType?.exhaustive === true,
		};
	}
	if (typeName.startsWith(PROJECT_TYPE_PREFIX)) {
		const projectType = projectClassMembersByName(ctx).get(
			typeName.slice(PROJECT_TYPE_PREFIX.length),
		);
		return projectType
			? {
				owner: projectType.name,
				members: projectType.members,
				exhaustive: projectType.exhaustive ?? projectType.kind === 'class',
			}
			: undefined;
	}
	const hostType = getHostType(typeName, ctx.model);
	return {
		owner: typeName,
		members: getHostMembers(typeName, ctx.model),
		exhaustive: hostType?.exhaustive === true,
	};
}

function resolveAnyMemberReturnType(
	ownerType: string,
	memberName: string,
	ctx: MemberCompletionContext,
): ResolvedMemberReturn | undefined {
	const union = parseUnionTypeKey(ownerType);
	if (union) {
		const resolved = union
			.map((item) => resolveAnyMemberReturnType(item, memberName, ctx))
			.filter((item): item is ResolvedMemberReturn => Boolean(item));
		if (resolved.length === 0) {
			return undefined;
		}
		return {
			type: typeKeyFor(resolved.map((item) => item.type)),
			kind: resolved.every((item) => item.kind === 'method') ? 'method' : 'property',
		};
	}
	const combined = parseCombinedTypeKey(ownerType);
	if (combined) {
		const projectType = projectClassMembersByName(ctx).get(combined.projectKey);
		const projectMember = projectType?.members.find(
			(m) => m.name.toLowerCase() === memberName.toLowerCase(),
		);
		if (projectMember?.returns) {
			const type = resolveDeclaredObjectType(projectMember.returns, ctx, ctx.model);
			return type ? { type, kind: projectMember.kind } : undefined;
		}
		return hostMemberReturn(combined.hostType, memberName, ctx.model);
	}
	if (!ownerType.startsWith(PROJECT_TYPE_PREFIX)) {
		return hostMemberReturn(ownerType, memberName, ctx.model);
	}
	const projectType = projectClassMembersByName(ctx).get(
		ownerType.slice(PROJECT_TYPE_PREFIX.length),
	);
	const member = projectType?.members.find(
		(m) => m.name.toLowerCase() === memberName.toLowerCase(),
	);
	if (!member?.returns) {
		return undefined;
	}
	const type = resolveDeclaredObjectType(member.returns, ctx, ctx.model);
	return type ? { type, kind: member.kind } : undefined;
}

function applyDefaultMemberReturnType(
	typeName: string | undefined,
	hasArguments: boolean,
	ctx: MemberCompletionContext,
): string | undefined {
	if (!typeName || !hasArguments) {
		return typeName;
	}
	const union = parseUnionTypeKey(typeName);
	if (union) {
		return typeKeyFor(
			union.map((item) => hostMemberReturn(item, 'Item', ctx.model)?.type ?? item),
		);
	}
	return hostMemberReturn(typeName, 'Item', ctx.model)?.type ?? typeName;
}

function hostMemberReturn(
	ownerType: string,
	memberName: string,
	model: HostObjectModel | undefined,
): ResolvedMemberReturn | undefined {
	const member = getHostMembers(ownerType, model).find(
		(m) => m.name.toLowerCase() === memberName.toLowerCase(),
	);
	if (member?.returns) {
		return { type: member.returns, kind: member.kind };
	}
	if (member?.returnsAnyOf?.length) {
		return { type: typeKeyFor(member.returnsAnyOf), kind: member.kind };
	}
	return undefined;
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
		const codeNameHost = ctx.codeNames?.[key];
		return codeNameHost ? combinedTypeKey(key, codeNameHost) : projectTypeKey(key);
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

function combinedTypeKey(projectKey: string, hostType: string): string {
	return `${COMBINED_TYPE_PREFIX}${projectKey}${COMBINED_TYPE_SEPARATOR}${hostType}`;
}

function parseCombinedTypeKey(
	typeName: string,
): { projectKey: string; hostType: string } | undefined {
	if (!typeName.startsWith(COMBINED_TYPE_PREFIX)) {
		return undefined;
	}
	const body = typeName.slice(COMBINED_TYPE_PREFIX.length);
	const sep = body.indexOf(COMBINED_TYPE_SEPARATOR);
	if (sep < 1 || sep >= body.length - 1) {
		return undefined;
	}
	return {
		projectKey: body.slice(0, sep),
		hostType: body.slice(sep + 1),
	};
}

function parseUnionTypeKey(typeName: string): string[] | undefined {
	if (!typeName.startsWith(UNION_TYPE_PREFIX)) {
		return undefined;
	}
	const parts = typeName
		.slice(UNION_TYPE_PREFIX.length)
		.split(UNION_TYPE_SEPARATOR)
		.filter((item) => item.length > 0);
	return parts.length > 0 ? parts : undefined;
}

function typeKeyFor(types: readonly string[]): string {
	const out: string[] = [];
	const seen = new Set<string>();
	for (const type of types) {
		for (const item of parseUnionTypeKey(type) ?? [type]) {
			const key = item.toLowerCase();
			if (seen.has(key)) {
				continue;
			}
			seen.add(key);
			out.push(item);
		}
	}
	return out.length === 1
		? out[0]
		: `${UNION_TYPE_PREFIX}${out.join(UNION_TYPE_SEPARATOR)}`;
}

function displayTypeName(typeName: string): string {
	const dot = typeName.lastIndexOf('.');
	return dot >= 0 ? typeName.slice(dot + 1) : typeName;
}

function mergeCompletionMembers(
	...memberGroups: readonly (readonly CompletionMemberSource[])[]
): CompletionMemberSource[] {
	const out: CompletionMemberSource[] = [];
	const seen = new Set<string>();
	for (const members of memberGroups) {
		for (const member of members) {
			const key = member.name.toLowerCase();
			if (seen.has(key)) {
				continue;
			}
			seen.add(key);
			out.push(member);
		}
	}
	return out;
}

function projectSourceSurfaceCompleteWhenMergedWithHost(
	projectType: VbaProjectClassMembers | undefined,
): boolean {
	if (!projectType) {
		return true;
	}
	if (projectType.kind === 'userform') {
		return false;
	}
	return true;
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

function isGenericObjectDeclaration(declaredType: string): boolean {
	const lower = simpleTypeName(declaredType)?.toLowerCase();
	return lower === 'object' || lower === 'variant';
}

function findSetAssignedObjectType(
	source: string,
	offset: number,
	name: string,
	ctx: MemberCompletionContext,
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
		const hit = latestSetAssignmentInBody(enclosing.body, source, offset, lower);
		if (hit) {
			return receiverTypeFromExpressionTokens(hit.valueTokens, source, hit.offset, ctx);
		}
	}

	let latest: SetAssignment | undefined;
	for (const member of module.members) {
		if (member.kind !== 'Statement' || member.span.end > offset) {
			continue;
		}
		const hit = setAssignment(source, member);
		if (hit?.name.toLowerCase() === lower) {
			latest = hit;
		}
	}
	return latest
		? receiverTypeFromExpressionTokens(latest.valueTokens, source, latest.offset, ctx)
		: undefined;
}

interface SetAssignment {
	name: string;
	valueTokens: VbaToken[];
	offset: number;
}

function latestSetAssignmentInBody(
	body: BodyNode[],
	source: string,
	offset: number,
	lowerName: string,
): SetAssignment | undefined {
	let latest: SetAssignment | undefined;
	for (const node of body) {
		if (node.kind === 'Statement') {
			if (node.span.end > offset) {
				continue;
			}
			const hit = setAssignment(source, node);
			if (hit?.name.toLowerCase() === lowerName) {
				latest = hit;
			}
		} else if ('body' in node && Array.isArray(node.body)) {
			const hit = latestSetAssignmentInBody(node.body, source, offset, lowerName);
			if (hit) {
				latest = hit;
			}
		}
	}
	return latest;
}

function setAssignment(source: string, stmt: StatementNode): SetAssignment | undefined {
	const tokens = tokenize(source.slice(stmt.span.start, stmt.span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	let i = 0;
	if (
		tokens.length >= 2 &&
		(tokens[0].kind === 'identifier' || tokens[0].kind === 'keyword') &&
		tokens[1].rawText === ':'
	) {
		i = 2;
	}
	if (tokens[i]?.rawText.toLowerCase() !== 'set') {
		return undefined;
	}
	const nameToken = tokens[i + 1];
	if (!nameToken || nameToken.kind !== 'identifier') {
		return undefined;
	}
	const equals = tokens[i + 2];
	if (!equals || equals.kind !== 'operator' || equals.rawText !== '=') {
		return undefined;
	}
	return {
		name: nameToken.rawText,
		valueTokens: tokens.slice(i + 3),
		offset: stmt.span.start,
	};
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
