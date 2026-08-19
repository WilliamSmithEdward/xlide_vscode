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
import { IDENT_RE, isIdentLike, statementTokensCached } from '../lexer/tokenHelpers';
import {
	MSFORMS_CONTROL_CLASS_NAMES,
	MSFORMS_REFERENCE_MEMBERS,
	type MsFormsMember,
} from '../host/msformsReferenceMembers';
import { VBA_USERFORM_EXTENDER_MEMBERS, VBA_USERFORM_TYPE } from '../host/userFormExtenderMembers';
import { completionCursorContext } from './cursorContext';
import { parseModule } from '../parser/parseModule';
import {
	BodyNode,
	isLeafStatement,
	LeafStatementNode,
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
	getHostType,
	resolveHostMemberSignature,
	resolveHostAlias,
	resolveHostGlobal,
} from '../host/hostModel';
import {
	resolveRuntimeObject,
	resolveRuntimeObjectType,
} from '../runtime/vbaRuntime';
import { hasDocContent, renderDocMarkdown } from '../docs/docModel';
import type { VbaDoc } from '../docs/docModel';
import type {
	VbaProjectClassMemberDefinition,
	VbaProjectClassMembers,
	VbaSymbolAttribute,
} from '../symbols/symbolModel';

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
	/** Project object type that `Me` resolves to in the current class/document module. */
	meProjectType?: string;
	/** Source-declared workbook object members and visible UDT fields, keyed by type. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/**
	 * Members the module has that its own text never declares - a UserForm's
	 * controls. Carries the type, so `RegionPick.` can offer ComboBox members
	 * rather than only escaping the undeclared-variable finding.
	 */
	implicitMembers?: readonly { name: string; type: string }[];
	/**
	 * True/default lets generic Object/Variant receivers narrow from preceding
	 * simple Set assignments. Hard diagnostics disable this because VBA still
	 * compile-binds those receivers late.
	 */
	allowSetAssignmentRefinement?: boolean;
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
	/**
	 * Pre-parsed AST of the analyzed source, when the caller already holds one
	 * (the diagnostics engine parses once per pass and resolves many member
	 * references). Used instead of re-parsing the full module for With-scan
	 * windows and declared-binding lookups. Must correspond exactly to the
	 * `source` string passed to the resolver alongside this context.
	 */
	parsedModule?: ModuleNode;
	/**
	 * Full-source significant tokens (comments removed, newlines kept), used to
	 * slice the prefix token stream by offset instead of re-lexing the source
	 * prefix per dotted reference. Only consulted when a token ends exactly at
	 * the requested offset; other offsets fall back to the prefix tokenizer.
	 * Must correspond exactly to the `source` string passed alongside.
	 */
	sourceTokens?: readonly VbaToken[];
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
	/** Markdown documentation rendered from source or host reference metadata. */
	documentation?: string;
	/** Raw documentation model, used by signature help for parameter notes. */
	doc?: VbaDoc;
	/** Source declaration locations, when this completion comes from workbook code. */
	definitions?: readonly VbaProjectClassMemberDefinition[];
	/** True when exported source marks this member as the VBA default member. */
	defaultMember?: boolean;
	/** Exported attribute lines attached to this member. */
	attributes?: readonly VbaSymbolAttribute[];
}

export interface ResolvedMemberSurface {
	owner: string;
	members: MemberCompletion[];
	exhaustive: boolean;
}

const PROJECT_TYPE_PREFIX = 'project:';
const COMBINED_TYPE_PREFIX = 'combined:';
const COMBINED_TYPE_SEPARATOR = '|';
const UNION_TYPE_PREFIX = 'union:';
const UNION_TYPE_SEPARATOR = '|';

type CompletionMemberSource = Pick<
	HostMember,
	'name' | 'kind' | 'returns' | 'signature' | 'doc'
> & {
	writable?: boolean;
	writeType?: string;
	definitions?: readonly VbaProjectClassMemberDefinition[];
	defaultMember?: boolean;
	attributes?: readonly VbaSymbolAttribute[];
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

/** A logical-line boundary: a newline or a statement-separating colon. */
function isBoundary(token: VbaToken): boolean {
	return token.kind === 'newline' || token.rawText === ':';
}

/**
 * Resolves the member completions available at `offset`. Returns an empty array
 * when the cursor is not in a member-access position or the receiver type
 * cannot be resolved to a known host or source-backed project type.
 */
export function resolveMemberCompletions(
	source: string,
	offset: number,
	ctx: MemberCompletionContext = {},
): MemberCompletion[] {
	const hit = memberSurfaceAtDot(source, offset, ctx);
	if (!hit) {
		return [];
	}
	const { currentType, surface, typedPrefix } = hit;
	const lowerPrefix = typedPrefix.toLowerCase();
	return surface.members
		.filter((mem) => mem.name.toLowerCase().startsWith(lowerPrefix))
		.map((mem) => completionFromSurfaceMember(currentType, surface, mem, ctx));
}

/**
 * Resolves the single member named `memberName` at `offset` without building
 * completion rows (and rendering documentation) for the whole member surface.
 */
export function resolveMemberCompletionNamed(
	source: string,
	offset: number,
	memberName: string,
	ctx: MemberCompletionContext = {},
): MemberCompletion | undefined {
	const hit = memberSurfaceAtDot(source, offset, ctx);
	if (!hit) {
		return undefined;
	}
	const lowerName = memberName.toLowerCase();
	const mem = hit.surface.members.find((m) => m.name.toLowerCase() === lowerName);
	return mem
		? completionFromSurfaceMember(hit.currentType, hit.surface, mem, ctx)
		: undefined;
}

/**
 * Resolves just the source definition locations of the member named
 * `memberName` ending at `offset`. Reference/rename providers call this once
 * per textual occurrence, so it skips completion-row construction (signature
 * lookup, documentation markdown) and bails on a cheap char-level scan when
 * no member-access dot precedes the name. Callers that already hold the
 * module's significant prefix tokens can pass them to skip the tokenization.
 */
export function resolveMemberDefinitionsAt(
	source: string,
	offset: number,
	memberName: string,
	ctx: MemberCompletionContext = {},
	prefixTokens?: VbaToken[],
): readonly VbaProjectClassMemberDefinition[] {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	if (!precededByMemberAccessDot(source, safeOffset - memberName.length)) {
		return [];
	}
	// Only trust supplied tokens that end exactly with the member name; when a
	// surrounding token swallows the name (e.g. a bracketed identifier), fall
	// back to tokenizing the prefix so behavior matches the unsliced path.
	const last = prefixTokens?.[prefixTokens.length - 1];
	const tokens =
		last && last.end === safeOffset &&
		last.rawText.toLowerCase() === memberName.toLowerCase()
			? prefixTokens
			: undefined;
	const hit = memberSurfaceAtDot(source, safeOffset, ctx, tokens);
	if (!hit) {
		return [];
	}
	const lowerName = memberName.toLowerCase();
	return hit.surface.members.find((m) => m.name.toLowerCase() === lowerName)
		?.definitions ?? [];
}

/**
 * Char-level fast path: true when the identifier starting at `nameStart` is
 * preceded by a member-access dot, allowing for whitespace and `_` line
 * continuations (the only trivia the lexer permits between the dot and the
 * member name).
 */
export function precededByMemberAccessDot(source: string, nameStart: number): boolean {
	let i = nameStart - 1;
	for (;;) {
		while (i >= 0 && (source[i] === ' ' || source[i] === '\t')) {
			i -= 1;
		}
		if (i < 0) {
			return false;
		}
		const ch = source[i];
		if (ch === '.') {
			return true;
		}
		if (ch === '\n' || ch === '\r') {
			if (ch === '\n' && i > 0 && source[i - 1] === '\r') {
				i -= 1;
			}
			i -= 1;
			while (i >= 0 && (source[i] === ' ' || source[i] === '\t')) {
				i -= 1;
			}
			if (i < 0 || source[i] !== '_') {
				return false;
			}
			i -= 1;
			continue;
		}
		return false;
	}
}

/**
 * Significant prefix tokens (comments removed, newlines kept) for `offset`.
 * When the context carries full-source tokens and a token ends exactly at
 * `offset`, slices the shared stream instead of re-lexing the prefix; the two
 * paths produce identical tokens because the cut sits on a token boundary.
 */
function prefixSignificantTokens(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): VbaToken[] {
	const shared = ctx.sourceTokens;
	if (shared && shared.length > 0) {
		// Binary search for the last token with end <= offset.
		let lo = 0;
		let hi = shared.length - 1;
		let found = -1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			if (shared[mid].end <= offset) {
				found = mid;
				lo = mid + 1;
			} else {
				hi = mid - 1;
			}
		}
		if (found >= 0 && shared[found].end === offset) {
			// Receiver chains never cross a logical-statement boundary and every
			// consumer walks backward stopping at one, so the prefix only needs
			// to reach back to the previous newline token (kept as the boundary
			// marker). Slicing from module start instead copies O(module) tokens
			// per query, which turns a large-module analysis pass quadratic.
			let start = found;
			while (start > 0 && shared[start].kind !== 'newline') {
				start--;
			}
			return shared.slice(start, found + 1);
		}
	}
	return completionCursorContext(source, offset).significantTokens;
}

function memberSurfaceAtDot(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
	prefixTokens?: VbaToken[],
): { currentType: string; surface: MemberSurface; typedPrefix: string } | undefined {
	// Keep newline tokens: they mark statement boundaries so a dangling
	// member-access dot on a previous line is not merged into this chain.
	const tokens = prefixTokens ?? prefixSignificantTokens(source, offset, ctx);
	if (tokens.length === 0) {
		return undefined;
	}

	// Identify the typed member prefix (text after the dot) and the dot itself.
	let i = tokens.length - 1;
	let typedPrefix = '';
	if (isIdentLike(tokens[i]) && i > 0 && tokens[i - 1].rawText === '.') {
		typedPrefix = tokens[i].rawText;
		i -= 1;
	}
	if (i < 0 || tokens[i].rawText !== '.') {
		return undefined;
	}
	// tokens[i] is the member-access dot; the receiver chain ends at i-1.
	const currentType = receiverTypeFromTokens(tokens, i, source, offset, ctx);
	if (!currentType) {
		return undefined;
	}
	const surface = memberSurfaceForType(currentType, ctx);
	if (!surface) {
		return undefined;
	}
	return { currentType, surface, typedPrefix };
}

/**
 * Resolves the complete source/host member surface at a member-access dot,
 * including empty-but-exhaustive project surfaces that cannot be represented by
 * completion rows alone.
 */
export function resolveMemberSurfaceAt(
	source: string,
	offset: number,
	ctx: MemberCompletionContext = {},
): ResolvedMemberSurface | undefined {
	const currentType = resolveReceiverTypeAt(source, offset, ctx);
	if (!currentType) {
		return undefined;
	}
	const surface = memberSurfaceForType(currentType, ctx);
	if (!surface) {
		return undefined;
	}
	return {
		owner: surface.owner,
		exhaustive: surface.exhaustive,
		members: surface.members.map((mem) =>
			completionFromSurfaceMember(currentType, surface, mem, ctx),
		),
	};
}

function completionFromSurfaceMember(
	currentType: string,
	surface: MemberSurface,
	mem: CompletionMemberSource,
	ctx: MemberCompletionContext,
): MemberCompletion {
	return {
		name: mem.name,
		kind: mem.kind,
		returns: mem.returns,
		signature: mem.signature ?? signatureForMember(currentType, mem.name, ctx),
		writable: mem.writable,
		writeType: mem.writeType,
		owner: surface.owner,
		surfaceExhaustive: surface.exhaustive,
		documentation: hasDocContent(mem.doc)
			? renderDocMarkdown(mem.doc)
			: undefined,
		doc: mem.doc,
		definitions: mem.definitions,
		defaultMember: mem.defaultMember,
		attributes: mem.attributes,
	};
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
		return (
			projectMemberSignature(combined.projectKey, memberName, ctx) ??
			resolveHostMemberSignature(combined.hostType, memberName, ctx.model)
		);
	}
	if (typeName.startsWith(PROJECT_TYPE_PREFIX)) {
		return projectMemberSignature(
			typeName.slice(PROJECT_TYPE_PREFIX.length),
			memberName,
			ctx,
		);
	}
	const runtimeObject = resolveRuntimeObjectType(typeName);
	if (runtimeObject) {
		return runtimeObject.members.find(
			(member) => member.name.toLowerCase() === memberName.toLowerCase(),
		)?.signature;
	}
	return resolveHostMemberSignature(typeName, memberName, ctx.model);
}

function projectMemberSignature(
	projectKey: string,
	memberName: string,
	ctx: MemberCompletionContext,
): string | undefined {
	const projectType = projectClassMembersByName(ctx).get(projectKey);
	return projectType?.members.find(
		(member) => member.name.toLowerCase() === memberName.toLowerCase(),
	)?.signature;
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
	const tokens = prefixSignificantTokens(source, offset, ctx);
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
	const explicitReceiver = receiverTypeFromChain(chain, source, offset, ctx);
	if (explicitReceiver) {
		return explicitReceiver;
	}
	const groupedReceiver = receiverTypeFromParenthesizedReceiver(
		tokens,
		dotIndex - 1,
		source,
		offset,
		ctx,
	);
	if (groupedReceiver) {
		return groupedReceiver;
	}
	const implicitWithChain = collectImplicitWithChain(tokens, dotIndex - 1);
	if (implicitWithChain === undefined) {
		return undefined;
	}
	return receiverTypeFromImplicitWithChain(
		withReceiverTypeAt(source, tokens[dotIndex].end, ctx),
		implicitWithChain,
		ctx,
	);
}

// Members whose declared return IS the already-resolved element/result: the
// default member (Item/_Default) and the creation method Add. A call to one of
// these must not be element-indexed again, or a collection whose element is
// itself a collection (e.g. SparklineGroups.Item(1)) over-resolves one level.
function isExplicitElementAccessor(name: string): boolean {
	const lower = name.toLowerCase();
	return lower === 'item' || lower === '_default' || lower === 'add';
}

function receiverTypeFromImplicitWithChain(
	withType: string | undefined,
	chain: ReceiverChainSegment[],
	ctx: MemberCompletionContext,
): string | undefined {
	let currentType = withType;
	for (const segment of chain) {
		if (!currentType) {
			return undefined;
		}
		const resolved = resolveAnyMemberReturnType(currentType, segment.name, ctx);
		if (!resolved) {
			return undefined;
		}
		// A member called with arguments indexes into its return type; when that
		// type is a host collection, applyDefaultMemberReturnType resolves the
		// element (and no-ops otherwise). This holds for method-kind accessors too
		// (e.g. ws.ChartObjects(1).Chart), so it must not be gated on kind. But
		// Item/_Default/Add already return the resolved element/result, so they are
		// not re-indexed (avoids over-resolving SparklineGroups.Item(1) one level).
		currentType = applyDefaultMemberReturnType(
			resolved.type,
			segment.hasArguments && !isExplicitElementAccessor(segment.name),
			ctx,
		);
	}
	return currentType;
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

function receiverTypeFromParenthesizedReceiver(
	tokens: VbaToken[],
	endIndex: number,
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	if (endIndex < 0 || tokens[endIndex].rawText !== ')') {
		return undefined;
	}
	const open = matchParenLeft(tokens, endIndex);
	if (open < 0) {
		return undefined;
	}
	const expressionTokens = tokens.slice(open + 1, endIndex);
	return (
		receiverTypeFromExpressionTokens(expressionTokens, source, offset, ctx) ??
		receiverTypeFromParenthesizedReceiver(
			expressionTokens,
			expressionTokens.length - 1,
			source,
			offset,
			ctx,
		)
	);
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
		// A member called with arguments indexes into its return type; when that
		// type is a host collection, applyDefaultMemberReturnType resolves the
		// element (and no-ops otherwise). This holds for method-kind accessors too
		// (e.g. ws.ChartObjects(1).Chart), so it must not be gated on kind. But
		// Item/_Default/Add already return the resolved element/result, so they are
		// not re-indexed (avoids over-resolving SparklineGroups.Item(1) one level).
		currentType = applyDefaultMemberReturnType(
			resolved.type,
			segment.hasArguments && !isExplicitElementAccessor(segment.name),
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
			// Empty parens foo() are a call with no index, not collection indexing;
			// only a non-empty argument list resolves to an element (matches the
			// assignment-inference path's argumentTokens.length > 0 check).
			if (open < i - 1) {
				pendingHasArguments = true;
			}
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

// `Me` is the only VBA keyword that can terminate a receiver expression (`Me.`);
// every other keyword before a dot (In, To, Then, ...) introduces a fresh
// expression, so the dot is a leading implicit-With member access.
const RECEIVER_TAIL_KEYWORDS = new Set(['me']);

/**
 * True when `token` (the token immediately before a `.`) means the dot is a
 * LEADING implicit-With member-access dot rather than `receiver.member`. A dot is
 * explicit only when preceded by something that terminates a receiver expression:
 * a plain identifier, `Me`, or a closing `)`/`]`. Anything else - a statement
 * boundary, an operator (`=`, `&`, `+`, ...), `(`/`,`, or an expression-introducing
 * keyword (`In`, `To`, `Then`, ...) - starts a new expression where `.member`
 * binds to the active `With` block (e.g. `For Each wb In .Workbooks`, `Set x = .Foo`).
 */
function precedesLeadingMemberDot(token: VbaToken): boolean {
	// A plain identifier or a foreign-name escape `[Foo]` (lexed as one
	// bracketedIdentifier token) terminates a receiver, so the following dot is an
	// explicit `receiver.member`, not an implicit-With leading dot.
	if (token.kind === 'identifier' || token.kind === 'bracketedIdentifier') {
		return false;
	}
	if (token.rawText === ')' || token.rawText === ']') {
		return false;
	}
	if (token.kind === 'keyword' && RECEIVER_TAIL_KEYWORDS.has(token.rawText.toLowerCase())) {
		return false;
	}
	return true;
}

function collectImplicitWithChain(
	tokens: VbaToken[],
	endIndex: number,
): ReceiverChainSegment[] | undefined {
	if (endIndex < 0 || precedesLeadingMemberDot(tokens[endIndex])) {
		return [];
	}
	const segments: ReceiverChainSegment[] = [];
	let i = endIndex;
	let pendingHasArguments = false;
	for (;;) {
		if (i >= 0 && isBoundary(tokens[i])) {
			return undefined;
		}
		if (i >= 0 && tokens[i].rawText === ')') {
			const open = matchParenLeft(tokens, i);
			if (open < 0) {
				return undefined;
			}
			// Empty parens foo() are a call with no index, not collection indexing;
			// only a non-empty argument list resolves to an element (matches the
			// assignment-inference path's argumentTokens.length > 0 check).
			if (open < i - 1) {
				pendingHasArguments = true;
			}
			i = open - 1;
			continue;
		}
		if (i < 0 || !isIdentLike(tokens[i])) {
			return undefined;
		}
		segments.unshift({
			name: word(tokens[i]),
			hasArguments: pendingHasArguments,
		});
		pendingHasArguments = false;
		i -= 1;
		if (i >= 0 && tokens[i].rawText === '.') {
			const prior = i - 1;
			if (prior < 0 || precedesLeadingMemberDot(tokens[prior])) {
				return segments;
			}
			i = prior;
			continue;
		}
		return undefined;
	}
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
		const projectKey = projectKeyForTypeName(ctx.meProjectType, ctx);
		if (ctx.meType) {
			return projectKey ? combinedTypeKey(projectKey, ctx.meType) : ctx.meType;
		}
		return projectKey ? projectTypeKey(projectKey) : undefined;
	}
	const declared = findDeclaredBinding(source, offset, root, ctx);
	if (declared) {
		if (declared.asType) {
			const declaredObjectType = resolveDeclaredObjectType(declared.asType, ctx, model);
			if (declaredObjectType) {
				return declaredObjectType;
			}
			if (!isGenericObjectDeclaration(declared.asType)) {
				return undefined;
			}
		}
		return ctx.allowSetAssignmentRefinement === false
			? undefined
			: findSetAssignedObjectType(source, offset, root, ctx);
	}
	const implicitMember = (ctx.implicitMembers ?? []).find(
		(member) => member.name.toLowerCase() === lower,
	);
	if (implicitMember) {
		return implicitMember.type;
	}
	const projectSurface = projectClassMembersByName(ctx).get(lower);
	const projectKey = projectSurface ? lower : undefined;
	const runtimeObject = resolveRuntimeObject(root);
	if (runtimeObject) {
		return runtimeObject.type;
	}
	const asGlobal = resolveHostGlobal(root, model);
	if (asGlobal) {
		return projectKey ? combinedTypeKey(projectKey, asGlobal) : asGlobal;
	}
	const asCode = ctx.codeNames?.[lower];
	if (asCode) {
		return projectKey ? combinedTypeKey(projectKey, asCode) : asCode;
	}
	if (
		projectSurface?.kind === 'standardModule'
		|| projectSurface?.kind === 'class'
		|| projectSurface?.kind === 'userform'
		// An Enum name reaches its constants: `Corner.TopLeft` is ordinary VBA,
		// and is how a reader tells one enum's TopLeft from another's.
		|| projectSurface?.kind === 'enum'
	) {
		// A standard module's name reaches its members. A class or UserForm name does too:
		// UserForms always carry their predeclared default instance, and factory-style classes
		// (VB_PredeclaredId) are addressed by name as a matter of course. The attribute itself
		// is invisible to a host that reads module text without its header, so the offer is
		// not gated on it; misusing a class that is not predeclared is the diagnostics' concern.
		return projectTypeKey(lower);
	}
	return ctx.allowSetAssignmentRefinement === false
		? undefined
		: findSetAssignedObjectType(source, offset, root, ctx);
}

function withReceiverTypeAt(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): string | undefined {
	let currentType: string | undefined;
	for (const expression of activeWithExpressionsAt(source, offset, ctx)) {
		const explicitType = receiverTypeFromExpressionTokens(
			expression.tokens,
			source,
			expression.sliceStart,
			ctx,
		);
		if (explicitType) {
			currentType = explicitType;
			continue;
		}
		const implicitChain = collectImplicitWithChain(
			expression.tokens,
			expression.tokens.length - 1,
		);
		if (implicitChain === undefined) {
			return undefined;
		}
		currentType = receiverTypeFromImplicitWithChain(currentType, implicitChain, ctx);
		if (!currentType) {
			return undefined;
		}
	}
	return currentType;
}

interface ActiveWithExpression {
	tokens: VbaToken[];
	sliceStart: number;
}

function activeWithExpressionsAt(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): ActiveWithExpression[] {
	const scan = activeWithScanWindow(source, offset, ctx);
	const stack: ActiveWithExpression[] = [];
	let statement: VbaToken[] = [];
	const flush = (): void => {
		processWithStackStatement(statement, stack, scan.sliceStart);
		statement = [];
	};
	for (const token of tokenize(scan.text)) {
		if (token.kind === 'comment') {
			continue;
		}
		if (isBoundary(token)) {
			flush();
			continue;
		}
		statement.push(token);
	}
	flush();
	return stack;
}

function activeWithScanWindow(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): { text: string; sliceStart: number } {
	const safeOffset = Math.max(0, offset);
	const module: ModuleNode = ctx.parsedModule ?? parseModule(source);
	const enclosing = module.members.find(
		(mem): mem is ProcedureNode =>
			mem.kind === 'Procedure' &&
			safeOffset >= mem.span.start &&
			safeOffset <= mem.span.end,
	);
	if (!enclosing) {
		return { text: source.slice(0, safeOffset), sliceStart: 0 };
	}
	return {
		text: source.slice(enclosing.span.start, safeOffset),
		sliceStart: enclosing.span.start,
	};
}

function processWithStackStatement(
	statement: readonly VbaToken[],
	stack: ActiveWithExpression[],
	sliceStart: number,
): void {
	const start = statementExecutableStart(statement);
	const first = statement[start];
	if (!first) {
		return;
	}
	const firstWord = word(first).toLowerCase();
	if (firstWord === 'with') {
		stack.push({
			tokens: statement.slice(start + 1),
			sliceStart: sliceStart + first.start,
		});
		return;
	}
	if (firstWord === 'end' && word(statement[start + 1] ?? first).toLowerCase() === 'with') {
		stack.pop();
	}
}

function statementExecutableStart(statement: readonly VbaToken[]): number {
	if (
		statement.length > 1 &&
		statement[0].kind === 'integerLiteral' &&
		/^\d+$/.test(statement[0].rawText)
	) {
		return 1;
	}
	if (
		statement.length > 2 &&
		isIdentLike(statement[0]) &&
		statement[1].rawText === ':'
	) {
		return 2;
	}
	return 0;
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
		const controls = implicitMembersOf(combined.projectKey, ctx);
		// A form's `Me` is combined:<form>|MSForms.UserForm, and MSForms is not
		// part of the Excel host model, so the base surface comes from the
		// forms metadata when the host type names a forms class.
		const formsMembers = msFormsControlMembers(combined.hostType);
		const baseMembers = formsMembers ?? getHostMembers(combined.hostType, ctx.model);
		if (!projectType && !hostType && !formsMembers && controls.length === 0) {
			return undefined;
		}
		return {
			owner: projectType?.name ?? combined.hostType,
			members: mergeCompletionMembers(projectType?.members ?? [], controls, baseMembers),
			// A form's own `Me` follows the same authority rule as its
			// qualified name (#26): the forms base plus an index-proven
			// control list proves absence. Other combined surfaces keep the
			// host-exhaustive gate.
			exhaustive: formsMembers
				? projectType?.exhaustive === true
				: controls.length === 0 &&
					projectSourceSurfaceCompleteWhenMergedWithHost(projectType) &&
					hostType?.exhaustive === true,
		};
	}
	if (typeName.startsWith(PROJECT_TYPE_PREFIX)) {
		const projectKey = typeName.slice(PROJECT_TYPE_PREFIX.length);
		const projectType = projectClassMembersByName(ctx).get(projectKey);
		const controls = implicitMembersOf(projectKey, ctx);
		if (!projectType) {
			return controls.length > 0
				? { owner: ctx.meProjectType ?? projectKey, members: controls, exhaustive: false }
				: undefined;
		}
		if (projectType.kind === 'userform') {
			// A form IS an MSForms.UserForm wherever it is reached from, so a
			// qualified reference from another module gets Show, Hide and the
			// rest of the form surface alongside the form's code and controls
			// (#22). Exhaustive exactly when the index proved the control list
			// (host-supplied or parsed from a `.frm` designer header): the
			// merged code-behind + controls + UserForm base surface then
			// proves absence the same way the VBE's compiler does (#26).
			return {
				owner: projectType.name,
				members: mergeCompletionMembers(
					projectType.members,
					controls,
					msFormsControlMembers(VBA_USERFORM_TYPE) ?? [],
				),
				exhaustive: projectType.exhaustive === true,
			};
		}
		return {
			owner: projectType.name,
			members: mergeCompletionMembers(projectType.members, controls),
			exhaustive: controls.length > 0
				? false
				: projectType.exhaustive ?? projectType.kind === 'class',
		};
	}
	const runtimeObject = resolveRuntimeObjectType(typeName);
	if (runtimeObject) {
		return {
			owner: runtimeObject.name,
			members: runtimeObject.members,
			exhaustive: runtimeObject.exhaustive,
		};
	}
	const controlMembers = msFormsControlMembers(typeName);
	if (controlMembers) {
		return {
			owner: typeName,
			members: controlMembers,
			// Not exhaustive: this list is for offering members, and treating it
			// as complete would let absence become a diagnostic about form code.
			exhaustive: false,
		};
	}
	const hostType = getHostType(typeName, ctx.model);
	return {
		owner: typeName,
		members: getHostMembers(typeName, ctx.model),
		exhaustive: hostType?.exhaustive === true,
	};
}

const MSFORMS_CONTROL_CLASS_SET: ReadonlySet<string> = new Set(MSFORMS_CONTROL_CLASS_NAMES);

/**
 * Members of `MSForms.ComboBox` and friends, for a form's controls - and of
 * `MSForms.UserForm` for the form itself, where VBA's own additions (Show,
 * Hide, Name, Left, ...) join the type library's list.
 *
 * A placed control also carries the `Control` base surface - Left, Top,
 * Visible, Name, SetFocus, Move, ZOrder - which the library declares once on
 * `MSForms.Control` rather than repeating per type, so it is merged here. The
 * per-type list wins where a name appears in both.
 */
export function msFormsControlMembers(typeName: string): HostMember[] | undefined {
	const match = /^MSForms\.([A-Za-z][\w]*)$/.exec(typeName);
	const reference = match ? MSFORMS_REFERENCE_MEMBERS[match[1]] : undefined;
	let members: readonly MsFormsMember[] | undefined = reference;
	if (typeName === VBA_USERFORM_TYPE) {
		members = [...VBA_USERFORM_EXTENDER_MEMBERS, ...(reference ?? [])];
	} else if (match && MSFORMS_CONTROL_CLASS_SET.has(match[1])) {
		const own = new Set((reference ?? []).map((member) => member.name.toLowerCase()));
		members = [
			...(reference ?? []),
			...(MSFORMS_REFERENCE_MEMBERS['Control'] ?? []).filter(
				(member) => !own.has(member.name.toLowerCase()),
			),
		];
	}
	if (!members || members.length === 0) {
		return undefined;
	}
	return members.map((member) => ({
		name: member.name,
		kind: member.kind,
		returns: member.returns,
		readOnly: member.readOnly,
		signature: member.signature,
	})) as HostMember[];
}

/**
 * A form's controls are members of the form itself, so `Me.` - and the form's
 * own name, which reaches its predeclared instance - offers them alongside the
 * code it declares. Only the module being edited has a control list to offer:
 * the context carries one, and it is that module's.
 */
function implicitMembersOf(
	projectKey: string,
	ctx: MemberCompletionContext,
): CompletionMemberSource[] {
	if (!ctx.implicitMembers?.length || ctx.meProjectType?.toLowerCase() !== projectKey) {
		return [];
	}
	return ctx.implicitMembers.map((member) => ({
		name: member.name,
		kind: 'property' as HostMemberKind,
		returns: member.type,
	}));
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
		return implicitMemberReturn(combined.projectKey, memberName, ctx)
			?? hostMemberReturn(combined.hostType, memberName, ctx.model);
	}
	if (!ownerType.startsWith(PROJECT_TYPE_PREFIX)) {
		const runtimeObject = resolveRuntimeObjectType(ownerType);
		if (runtimeObject) {
			const member = runtimeObject.members.find(
				(item) => item.name.toLowerCase() === memberName.toLowerCase(),
			);
			return member?.returns ? { type: member.returns, kind: member.kind } : undefined;
		}
		return hostMemberReturn(ownerType, memberName, ctx.model);
	}
	const projectKey = ownerType.slice(PROJECT_TYPE_PREFIX.length);
	const projectType = projectClassMembersByName(ctx).get(projectKey);
	const member = projectType?.members.find(
		(m) => m.name.toLowerCase() === memberName.toLowerCase(),
	);
	if (!member?.returns) {
		return implicitMemberReturn(projectKey, memberName, ctx);
	}
	const type = resolveDeclaredObjectType(member.returns, ctx, ctx.model);
	return type ? { type, kind: member.kind } : undefined;
}

/** `Me.RegionPick.` chains through the control's own type. */
function implicitMemberReturn(
	projectKey: string,
	memberName: string,
	ctx: MemberCompletionContext,
): ResolvedMemberReturn | undefined {
	const control = implicitMembersOf(projectKey, ctx).find(
		(member) => member.name.toLowerCase() === memberName.toLowerCase(),
	);
	return control?.returns ? { type: control.returns, kind: 'property' } : undefined;
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
	// The project's own declarations are consulted BEFORE the referenced type
	// libraries, which is what VBA does. The Excel object model owns a lot of
	// ordinary nouns - Point, Border, Font, Shape, Style, Name - so a developer
	// who declares one of those got the library type's members instead of their
	// own, with nothing to indicate the name was ambiguous.
	const key = projectKeyForTypeName(declaredType, ctx);
	if (key) {
		const codeNameHost = ctx.codeNames?.[key];
		return codeNameHost ? combinedTypeKey(key, codeNameHost) : projectTypeKey(key);
	}
	return resolveHostAlias(declaredType, model) ?? resolveMsFormsTypeName(declaredType);
}

/**
 * Canonical `MSForms.<Type>` for a declared type the forms metadata knows,
 * case-insensitively - `Dim t As MSForms.TextBox` and a control member typed
 * `MSForms.ComboBox` both chain through it. Qualified names only: a bare
 * `TextBox` stays unresolved rather than guessed, since without the reference
 * line we cannot know MSForms is what it means.
 */
function resolveMsFormsTypeName(declaredType: string): string | undefined {
	const match = /^MSForms\s*\.\s*([A-Za-z][\w]*)$/i.exec(declaredType.trim());
	if (!match) {
		return undefined;
	}
	const lower = match[1].toLowerCase();
	for (const key of Object.keys(MSFORMS_REFERENCE_MEMBERS)) {
		if (key.toLowerCase() === lower) {
			return `MSForms.${key}`;
		}
	}
	return undefined;
}

function projectKeyForTypeName(
	typeName: string | undefined,
	ctx: MemberCompletionContext,
): string | undefined {
	if (!typeName) {
		return undefined;
	}
	const key = simpleTypeName(typeName)?.toLowerCase();
	if (!key) {
		return undefined;
	}
	const projectType = projectClassMembersByName(ctx).get(key);
	// A standard module is not a type you can declare against, and an Enum is a
	// VALUE type: `Dim c As Corner` is a Long, not an object. Both are member
	// surfaces so `Module.Member` and `Corner.TopLeft` resolve, but neither may
	// answer here or a plain enum variable would look like an object.
	return projectType
		&& projectType.kind !== 'standardModule'
		&& projectType.kind !== 'enum'
		? key
		: undefined;
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
	const module: ModuleNode = ctx.parsedModule ?? parseModule(source);
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
		if (isLeafStatement(node)) {
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

function setAssignment(source: string, stmt: LeafStatementNode): SetAssignment | undefined {
	const tokens = statementTokensCached(source, stmt.span);
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
 * Finds a local variable, parameter, or module-level variable named `name`,
 * preferring the declaration in the procedure that encloses `offset`. Untyped
 * declarations still shadow globals, so callers need to know about them even
 * when there is no raw `As` type text.
 */
interface DeclaredBinding {
	asType?: string;
}

function findDeclaredBinding(
	source: string,
	offset: number,
	name: string,
	ctx: MemberCompletionContext,
): DeclaredBinding | undefined {
	const module: ModuleNode = ctx.parsedModule ?? parseModule(source);
	const lower = name.toLowerCase();

	const enclosing = module.members.find(
		(mem): mem is ProcedureNode =>
			mem.kind === 'Procedure' &&
			offset >= mem.span.start &&
			offset <= mem.span.end,
	);

	if (enclosing) {
		for (const param of enclosing.params) {
			if (param.name.toLowerCase() === lower) {
				return { asType: param.asType };
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
function findInBody(body: BodyNode[], lower: string): DeclaredBinding | undefined {
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

function matchGroup(group: VariableGroupNode, lower: string): DeclaredBinding | undefined {
	for (const decl of group.declarations) {
		if (decl.name.toLowerCase() === lower) {
			return { asType: decl.asType };
		}
	}
	return undefined;
}
