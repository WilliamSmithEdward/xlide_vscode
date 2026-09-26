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
	getHostEnumMembers,
	getHostType,
	resolveHostMemberSignature,
	resolveHostAlias,
	resolveHostGlobal,
	resolveHostEnum,
	resolveHostGlobalMember,
	bareTypeName,
} from '../host/hostModel';
import { derivedConstantDoc, derivedMemberDoc } from '../host/hostMemberDocs';
import { hostTypeResolvesWhenCompiling } from '../host/typeExtensibility';
import {
	resolveRuntimeObject,
	resolveRuntimeObjectType,
	resolveVbaLibraryQualifier,
} from '../runtime/vbaRuntime';
import { hasDocContent, renderDocMarkdown } from '../docs/docModel';
import type { VbaDoc } from '../docs/docModel';
import {
	isDataBoundDesignerClass,
	type VbaProjectClassMemberDefinition,
	type VbaProjectClassMembers,
	type VbaSymbolAttribute,
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
	/** Source-declared project object members and visible UDT fields, keyed by type. */
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
	/**
	 * Per-pass memo of the active `With` stack, keyed by enclosing procedure
	 * start. Callers resolving many references against one immutable source pass
	 * a fresh Map; the scan is then paid once per procedure rather than once per
	 * leading-dot member. Must be discarded whenever `source` changes.
	 */
	withScanCache?: Map<number, WithScanIndex>;
	/**
	 * Receiver-chain prefix results for one analysis pass, keyed by the chain's
	 * root token offset and the number of segments resolved (issue #135). The
	 * caller owns its lifetime: one source, one pass.
	 */
	receiverTypeCache?: Map<string, string | undefined>;
	/** Receiver chains already collected, keyed by the dot token's offset (issue #135). */
	receiverChainCache?: Map<number, ReceiverChain>;
}

/** A single member-completion result. */
export interface MemberCompletion {
	name: string;
	kind: HostMemberKind;
	/** Qualified type the member returns, when chainable. */
	returns?: string;
	/** Verified call signature, when the host metadata has one. */
	signature?: string;
	/** The type a host property declares, when it is not a chainable object. */
	declaredType?: string;
	/** The read/write contract the type library states for a host property. */
	access?: HostMember['access'];
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
	/** Source declaration locations, when this completion comes from project code. */
	definitions?: readonly VbaProjectClassMemberDefinition[];
	/** True when exported source marks this member as the VBA default member. */
	defaultMember?: boolean;
	/** The setters a project property declares (issue #107); absent for host members. */
	letAccessor?: boolean;
	setAccessor?: boolean;
	/** Exported attribute lines attached to this member. */
	attributes?: readonly VbaSymbolAttribute[];
}

export interface ResolvedMemberSurface {
	owner: string;
	members: MemberCompletion[];
	exhaustive: boolean;
}

const PROJECT_TYPE_PREFIX = 'project:';
/** Receiver key for a host enumeration used as a qualifier: `XlAxisType.xlCategory`. */
const HOST_ENUM_PREFIX = 'hostEnum:';
const VBA_LIBRARY_PREFIX = 'vbaLibrary:';
const COMBINED_TYPE_PREFIX = 'combined:';
const COMBINED_TYPE_SEPARATOR = '|';
const UNION_TYPE_PREFIX = 'union:';
const UNION_TYPE_SEPARATOR = '|';

type CompletionMemberSource = Pick<
	HostMember,
	'name' | 'kind' | 'returns' | 'signature' | 'declaredType' | 'access' | 'doc' | 'hidden'
> & {
	writable?: boolean;
	writeType?: string;
	definitions?: readonly VbaProjectClassMemberDefinition[];
	defaultMember?: boolean;
	letAccessor?: boolean;
	setAccessor?: boolean;
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

export interface ReceiverChain {
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
		// OFFERED is narrower than KNOWN (issue #56): a member the type
		// library marks hidden or restricted, or whose name VBA cannot write
		// at all, is still resolved, hovered and coloured - code that names
		// one is real code - but it is never proposed, because accepting the
		// proposal would produce something the VBE refuses to compile.
		.filter((mem) => !mem.hidden)
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
 * The kind of the HOST member named `memberName` ending at `offset`, or
 * undefined when the receiver is not a host object or carries no such member.
 *
 * Semantic-token painting calls this once per member-access dot in a module, so
 * it stops at the member's kind rather than building a completion row. The
 * receiver is resolved exactly the way hover resolves it, which is the point:
 * what the editor can describe is what it should be willing to color.
 */
export function resolveHostMemberKindAt(
	source: string,
	offset: number,
	memberName: string,
	ctx: MemberCompletionContext = {},
): HostMemberKind | undefined {
	const hit = memberSurfaceAtDot(source, offset, ctx);
	if (!hit || !hostReceiverTypesOf(hit.currentType).some((type) => getHostType(type, ctx.model))) {
		return undefined;
	}
	const lowerName = memberName.toLowerCase();
	return hit.surface.members.find((member) => member.name.toLowerCase() === lowerName)?.kind;
}

/**
 * The host types a receiver key denotes, for deciding whether a member came
 * from a host library at all.
 *
 * A document module's code name resolves to a COMBINED key - the module's own
 * source surface joined to `Excel.Worksheet` - and the surface built from one
 * is owned by the project name. Reading the owner alone therefore said "not a
 * host type" and `Sheet1.Calculate` stopped painting beside `Me.Calculate`,
 * which is the pair issue #31 exists to keep identical (issue #44).
 */
function hostReceiverTypesOf(receiverType: string): string[] {
	const union = parseUnionTypeKey(receiverType);
	if (union) {
		return union.flatMap((item) => hostReceiverTypesOf(item));
	}
	const combined = parseCombinedTypeKey(receiverType);
	if (combined) {
		return [combined.hostType];
	}
	return receiverType.startsWith(PROJECT_TYPE_PREFIX) ? [] : [receiverType];
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
	// A host member the reference never described still gets a tooltip, composed
	// from what the type library declares and marked `derived`.
	const doc = mem.doc?.summary
		? mem.doc
		: (getHostType(surface.owner, ctx.model)
			? derivedMemberDoc(mem, surface.owner) ?? mem.doc
			: mem.doc);
	return {
		name: mem.name,
		kind: mem.kind,
		returns: mem.returns,
		signature: mem.signature ?? signatureForMember(currentType, mem.name, ctx),
		declaredType: mem.declaredType,
		access: mem.access,
		writable: mem.writable,
		writeType: mem.writeType,
		owner: surface.owner,
		surfaceExhaustive: surface.exhaustive,
		documentation: hasDocContent(doc)
			? renderDocMarkdown(doc)
			: undefined,
		doc,
		definitions: mem.definitions,
		defaultMember: mem.defaultMember,
		letAccessor: mem.letAccessor,
		setAccessor: mem.setAccessor,
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
	// A dot whose chain is the previous dot's plus one member takes that dot's
	// chain and adds the member, instead of walking the whole chain back again
	// (issue #135: a 4,000-member chain took 2.4 s, each dot re-walking it).
	const chain = chainExtendedFromPreviousDot(tokens, dotIndex, ctx) ?? collectReceiverChainWithStart(tokens, dotIndex - 1);
	if (chain && ctx.receiverChainCache) {
		ctx.receiverChainCache.set(tokens[dotIndex].start, chain);
	}
	// The chain's root resolves the same at every dot of one statement, so the
	// type walk resumes from the longest prefix already resolved.
	const cacheBase = chain && ctx.receiverTypeCache ? tokens[chain.startIndex]?.start : undefined;
	const explicitReceiver = receiverTypeFromChain(chain?.segments ?? [], source, offset, ctx, cacheBase);
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
export function isExplicitElementAccessor(name: string): boolean {
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
	cacheBase?: number,
): string | undefined {
	if (chain.length === 0) {
		return undefined;
	}
	// Prefix results of this chain, keyed by the root token's offset and the
	// number of segments resolved: the longest cached prefix is where the walk
	// resumes, and every prefix reached is stored for the next dot.
	const cache = cacheBase === undefined ? undefined : ctx.receiverTypeCache;
	const keyFor = (segments: number): string => `${cacheBase}:${segments}`;
	let resumeAt = 0;
	let currentType: string | undefined;
	if (cache) {
		for (let s = chain.length; s >= 1; s -= 1) {
			const key = keyFor(s);
			if (cache.has(key)) {
				currentType = cache.get(key);
				resumeAt = s;
				break;
			}
		}
	}
	if (resumeAt === 0) {
		const root = chain[0];
		const rootType = resolveRoot(root.name, source, offset, ctx);
		if (!rootType) {
			cache?.set(keyFor(1), undefined);
			return undefined;
		}
		currentType = applyDefaultMemberReturnType(rootType, root.hasArguments, ctx);
		cache?.set(keyFor(1), currentType);
		resumeAt = 1;
	}
	for (let s = resumeAt; s < chain.length && currentType; s += 1) {
		const segment = chain[s];
		const resolved = resolveAnyMemberReturnType(currentType, segment.name, ctx);
		if (!resolved) {
			cache?.set(keyFor(s + 1), undefined);
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
		cache?.set(keyFor(s + 1), currentType);
	}
	return currentType;
}

/**
 * The chain for the dot at `dotIndex` when the token before it is a plain
 * member name that follows an already-resolved dot: that dot's cached chain
 * plus this member. Any other shape (arguments, a root, a boundary) is
 * collected the long way.
 */
function chainExtendedFromPreviousDot(
	tokens: VbaToken[],
	dotIndex: number,
	ctx: MemberCompletionContext,
): ReceiverChain | undefined {
	const cache = ctx.receiverChainCache;
	const member = tokens[dotIndex - 1];
	const previousDot = tokens[dotIndex - 2];
	if (!cache || !member || !isIdentLike(member) || previousDot?.rawText !== '.') {
		return undefined;
	}
	const previous = cache.get(previousDot.start);
	if (!previous) {
		return undefined;
	}
	return {
		segments: [...previous.segments, { name: word(member), hasArguments: false }],
		startIndex: previous.startIndex,
	};
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
	// A member of the host's hidden Global interface with a typed return is a
	// receiver too: Excel's Union(a, b) yields a Range, Word's RecentFiles a
	// RecentFiles (issue #34). Ranked with the other host-injected names.
	const asGlobalMember = resolveHostGlobalMember(root, model)?.returns;
	if (asGlobalMember) {
		return projectKey ? combinedTypeKey(projectKey, asGlobalMember) : asGlobalMember;
	}
	// VBA's own enums and modules of constants reach their constants too:
	// `VbMsgBoxResult.vbYes`, `ColorConstants.vbRed`. VBA is first in every
	// project's references, so it has a name a host shares - `Constants.vbCrLf`
	// in Excel is VBA's module, not Excel's Constants enum.
	const asVbaLibrary = !projectKey ? resolveVbaLibraryQualifier(root) : undefined;
	if (asVbaLibrary?.constants) {
		return `${VBA_LIBRARY_PREFIX}${asVbaLibrary.name}`;
	}
	// An enum name reaches its own constants: `XlAxisType.xlCategory` is ordinary
	// VBA and is how a reader tells one library's xlNone from another's.
	const asEnum = !projectKey ? resolveHostEnum(root, model) : undefined;
	if (asEnum) {
		return `${HOST_ENUM_PREFIX}${asEnum.displayName}`;
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
	const index = withScanIndex(scan, ctx);
	// Resume from the last complete statement before `offset` rather than from
	// the top of the procedure, then finish the partial statement the offset
	// sits in. Same answer, paid once per procedure instead of once per dot.
	const at = lastBoundaryAtOrBefore(index.boundaries, offset);
	const stack = at < 0 ? [] : index.stacks[at].slice();
	let statement: VbaToken[] = [];
	const flush = (): void => {
		processWithStackStatement(statement, stack, index.sliceStart);
		statement = [];
	};
	for (let i = at < 0 ? 0 : index.resumeAt[at]; i < index.tokens.length; i += 1) {
		const token = index.tokens[i];
		// Boundaries are absolute; the fallback lexer numbers its tokens from the
		// start of the sliced procedure, so the window's own start is added back.
		if (token.end + index.sliceStart > offset) {
			break;
		}
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

/**
 * The active `With` stack after every complete statement of one procedure.
 *
 * Walking the procedure from its start for each leading-dot member is quadratic
 * in the procedure's length: a single procedure holding 1,200 `With` blocks cost
 * 433 ms to paint, against 78 ms for a module of the same size split into
 * ordinary procedures. Callers that resolve many references against one source
 * pass a `withScanCache`, and then each procedure is walked once.
 */
interface WithScanIndex {
	tokens: readonly VbaToken[];
	sliceStart: number;
	/** End offset of each complete statement, ascending. */
	boundaries: number[];
	/** Stack after the statement ending at the same position in `boundaries`. */
	stacks: ActiveWithExpression[][];
	/** Token index to resume scanning from, per boundary. */
	resumeAt: number[];
}

function withScanIndex(
	scan: { text: string; sliceStart: number; procedureStart: number; windowEnd: number },
	ctx: MemberCompletionContext,
): WithScanIndex {
	const cached = scan.procedureStart >= 0
		? ctx.withScanCache?.get(scan.procedureStart)
		: undefined;
	if (cached) {
		return cached;
	}
	const window = withScanTokens(scan, Number.MAX_SAFE_INTEGER, ctx);
	const boundaries: number[] = [];
	const stacks: ActiveWithExpression[][] = [];
	const resumeAt: number[] = [];
	const stack: ActiveWithExpression[] = [];
	let statement: VbaToken[] = [];
	for (let i = 0; i < window.tokens.length; i += 1) {
		const token = window.tokens[i];
		if (token.kind === 'comment') {
			continue;
		}
		if (!isBoundary(token)) {
			statement.push(token);
			continue;
		}
		processWithStackStatement(statement, stack, window.sliceStart);
		statement = [];
		boundaries.push(token.end + window.sliceStart);
		stacks.push(stack.slice());
		resumeAt.push(i + 1);
	}
	const index: WithScanIndex = {
		tokens: window.tokens,
		sliceStart: window.sliceStart,
		boundaries,
		stacks,
		resumeAt,
	};
	if (scan.procedureStart >= 0) {
		ctx.withScanCache?.set(scan.procedureStart, index);
	}
	return index;
}

/** Index of the last boundary at or before `offset`, or -1. */
function lastBoundaryAtOrBefore(boundaries: readonly number[], offset: number): number {
	let lo = 0;
	let hi = boundaries.length - 1;
	let found = -1;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		if (boundaries[mid] <= offset) {
			found = mid;
			lo = mid + 1;
		} else {
			hi = mid - 1;
		}
	}
	return found;
}

/**
 * Tokens of the active `With` scan window.
 *
 * Re-lexing the enclosing procedure for every leading-dot member is quadratic in
 * the procedure's length, and semantic-token painting asks once per dot: a
 * module of 400 `With` blocks cost 455 microseconds per dot against 2 for every
 * other receiver shape. When the caller holds the full-source stream, the window
 * is a slice of it - the tokens are then already at absolute offsets, so the
 * slice start the callers add is zero.
 */
function withScanTokens(
	scan: { text: string; sliceStart: number; windowEnd: number },
	offset: number,
	ctx: MemberCompletionContext,
): { tokens: readonly VbaToken[]; sliceStart: number } {
	const shared = ctx.sourceTokens;
	if (!shared || shared.length === 0) {
		return { tokens: [...tokenize(scan.text)], sliceStart: scan.sliceStart };
	}
	let lo = 0;
	let hi = shared.length;
	while (lo < hi) {
		const mid = (lo + hi) >> 1;
		if (shared[mid].start < scan.sliceStart) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	const limit = Math.min(offset, scan.windowEnd);
	let end = lo;
	while (end < shared.length && shared[end].end <= limit) {
		end += 1;
	}
	return { tokens: shared.slice(lo, end), sliceStart: 0 };
}

function activeWithScanWindow(
	source: string,
	offset: number,
	ctx: MemberCompletionContext,
): { text: string; sliceStart: number; procedureStart: number; windowEnd: number } {
	const safeOffset = Math.max(0, offset);
	const module: ModuleNode = ctx.parsedModule ?? parseModule(source);
	const enclosing = module.members.find(
		(mem): mem is ProcedureNode =>
			mem.kind === 'Procedure' &&
			safeOffset >= mem.span.start &&
			safeOffset <= mem.span.end,
	);
	if (!enclosing) {
		// Module level: the window is everything before the offset, and there is
		// no procedure to key an index on.
		return { text: source.slice(0, safeOffset), sliceStart: 0, procedureStart: -1, windowEnd: safeOffset };
	}
	return {
		text: source.slice(enclosing.span.start, enclosing.span.end),
		sliceStart: enclosing.span.start,
		procedureStart: enclosing.span.start,
		windowEnd: enclosing.span.end,
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
			owner: union.map(bareTypeName).join(' | '),
			members: mergeCompletionMembers(...surfaces.map((surface) => surface.members)),
			// A union is what the library declares Object - ActiveSheet, a
			// Sheets item - so VBA binds its members when it runs and a name on
			// none of the parts is not a compile error (issue #114:
			// `ActiveSheet.asdf` compiles). Never exhaustive.
			exhaustive: false,
		};
	}
	if (typeName.startsWith(VBA_LIBRARY_PREFIX)) {
		const qualifier = resolveVbaLibraryQualifier(typeName.slice(VBA_LIBRARY_PREFIX.length));
		if (!qualifier?.constants) {
			return undefined;
		}
		return {
			owner: qualifier.name,
			members: qualifier.constants.map((constant) => ({
				name: constant.name,
				kind: 'property' as const,
				declaredType: constant.type ?? qualifier.name,
				access: 'read-only' as const,
			})),
			// Its members are exactly the type library's, so one can be proved absent.
			exhaustive: true,
		};
	}
	if (typeName.startsWith(HOST_ENUM_PREFIX)) {
		const enumName = typeName.slice(HOST_ENUM_PREFIX.length);
		const constants = getHostEnumMembers(enumName, ctx.model);
		if (constants.length === 0) {
			return undefined;
		}
		return {
			owner: enumName,
			members: constants.map((constant) => ({
				name: constant.name,
				kind: 'property' as const,
				declaredType: enumName,
				access: 'read-only' as const,
				doc: constant.doc ?? derivedConstantDoc(constant, resolveHostEnum(enumName, ctx.model)),
			})),
			// An enum's members are exactly its constants, so this surface can
			// prove one absent - unlike the object types, which never can.
			exhaustive: true,
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
		if (projectType.kind === 'userform' && isDataBoundDesignerClass(projectType.designerClass)) {
			// An Access form or report is its own library's class, not a
			// UserForm: `Form_Orders.Requery` reaches Access.Form's members,
			// and Show and Hide are not among them. Never exhaustive - its
			// record-source fields are members no list here can name.
			return {
				owner: projectType.name,
				members: mergeCompletionMembers(
					projectType.members,
					controls,
					getHostMembers(projectType.designerClass as string, ctx.model),
				),
				exhaustive: false,
			};
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
		// A complete member list proves absence only where the type library
		// says VBA resolves against the interface while compiling. Most of
		// Excel's object model is extensible, so `Application.Match` - a
		// worksheet function on no interface at all - is ordinary VBA, and
		// calling it absent reported working code as an error.
		exhaustive: hostType?.exhaustive === true && hostTypeResolvesWhenCompiling(typeName),
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
			?? msFormsMemberReturn(combined.hostType, memberName)
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
		return msFormsMemberReturn(ownerType, memberName)
			?? hostMemberReturn(ownerType, memberName, ctx.model);
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

/**
 * `Views.SelectedItem.` chains into the returned object's own MSForms surface
 * (issue #32): the member's bare return name ("Tab", "Font") resolves to its
 * qualified type exactly when the forms metadata carries that surface, so a
 * primitive or unmodelled return ends the chain instead of guessing.
 */
function msFormsMemberReturn(
	ownerType: string,
	memberName: string,
): ResolvedMemberReturn | undefined {
	const member = msFormsControlMembers(ownerType)?.find(
		(candidate) => candidate.name.toLowerCase() === memberName.toLowerCase(),
	);
	if (!member?.returns) {
		return undefined;
	}
	const type = resolveMsFormsTypeName(`MSForms.${member.returns}`);
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
	const members = getHostMembers(ownerType, model);
	const member = members.find(
		(m) => m.name.toLowerCase() === memberName.toLowerCase(),
	);
	if (member?.returns) {
		// An accessor the library declares `As Object` is late bound however
		// well the model knows its element: `Worksheets(1).NoSuchMember`
		// compiles (issue #114). A one-part union keeps the element for
		// completion and chaining without closing its surface. The
		// hand-written collections carry the repaired type on Item, so the
		// library's word is read off `_Default` as well.
		const defaultMember = memberName.toLowerCase() === 'item'
			? members.find((m) => m.name === '_Default')
			: undefined;
		const declaredObject = [member, defaultMember].some((m) =>
			m?.declaredType === 'Object' || /\bAs Object\s*$/i.test(m?.signature ?? ''));
		return {
			type: declaredObject ? `${UNION_TYPE_PREFIX}${member.returns}` : member.returns,
			kind: member.kind,
		};
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
	let lateBound = false;
	for (const type of types) {
		const parts = parseUnionTypeKey(type);
		if (parts) {
			lateBound = true;
		}
		for (const item of parts ?? [type]) {
			const key = item.toLowerCase();
			if (seen.has(key)) {
				continue;
			}
			seen.add(key);
			out.push(item);
		}
	}
	// A one-part union stays a union: it marks a value the library declares
	// Object, whose members bind when it runs (issue #114).
	return out.length === 1 && !lateBound
		? out[0]
		: `${UNION_TYPE_PREFIX}${out.join(UNION_TYPE_SEPARATOR)}`;
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
	return moduleProcedureBinding(module, lower);
}

/**
 * A module-level procedure of this name, as a receiver.
 *
 * The module's own members shadow the host's globals, and this is where the
 * two used to disagree: a module VARIABLE named `rows` resolved from the
 * declaration above, while `Public Property Get rows() As Widget` fell through
 * to Excel's global `Rows`, so `rows.Where(p)` was measured against
 * `Excel.Range` and reported member-not-found on legal code (issue #68). The
 * names that collide are the ones every workbook uses - rows, columns, cells,
 * selection, names, sheets, application.
 *
 * A Function or Property Get yields its return type. A Sub, or a Property with
 * only Let/Set, yields nothing readable - but it still shadows the global, so
 * it binds with no type rather than letting the host answer for it.
 */
function moduleProcedureBinding(module: ModuleNode, lower: string): DeclaredBinding | undefined {
	let shadow: DeclaredBinding | undefined;
	for (const mem of module.members) {
		if (mem.kind !== 'Procedure' || mem.name.toLowerCase() !== lower) {
			continue;
		}
		if (mem.procKind === 'Function' || mem.procKind === 'PropertyGet') {
			return mem.returnType ? { asType: mem.returnType } : {};
		}
		shadow = {};
	}
	return shadow;
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
