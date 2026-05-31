// Hover resolver.
//
// Given VBA source and a cursor offset, this describes the identifier under the
// cursor: its declaration signature, kind, origin module, and visibility. It
// resolves user symbols from the live module's symbol graph and host symbols
// (globals, code names, and `receiver.member` members) from the host model.
//
// Pure analyzer code: depends only on the lexer, symbol graph, and host model,
// never on vscode. See docs/xlide_vba_language_service_roadmap.md (Phase 7).

import { tokenize } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	ModuleSymbolKind,
	VbaSymbol,
	isProcedureKind,
} from '../symbols/symbolModel';
import { Span } from '../parser/nodes';
import { HostObjectModel } from '../host/excelObjectModel';
import { getHostType, resolveHostGlobal } from '../host/hostModel';
import { resolveRuntimeFunction } from '../runtime/vbaRuntime';
import {
	MemberCompletion,
	resolveMemberCompletions,
} from '../completion/memberAccess';
import { DocRegistry } from '../docs/docRegistry';
import { VbaDoc, hasDocContent, renderDocMarkdown } from '../docs/docModel';
import type { VbaProjectClassMembers } from '../symbols/symbolModel';

/** A resolved hover description for the identifier under the cursor. */
export interface HoverInfo {
	/** Declaration signature line, rendered by the provider as VBA code. */
	signature: string;
	/** Plain-text detail lines (origin module, visibility, source note). */
	details: string[];
	/** Span of the identifier the hover applies to. */
	span: Span;
	/** Markdown documentation (summary, params, returns) when available. */
	documentation?: string;
}

/** Project/module facts the hover resolver needs from outside the source. */
export interface HoverContext {
	/** Lowercased worksheet/document code name -> qualified host type. */
	codeNames?: Record<string, string>;
	/** Qualified host type that `Me` resolves to in the current module. */
	meType?: string;
	/** Name of the module being edited (for "Declared in Module" text). */
	moduleName?: string;
	/** Kind of the module being edited. */
	moduleKind?: ModuleSymbolKind;
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
	/** Source-declared workbook class/UserForm/document members, keyed by type. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Developer-defined external documentation (overrides the curated library). */
	docRegistry?: DocRegistry;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

function contains(span: Span, offset: number): boolean {
	return offset >= span.start && offset <= span.end;
}

/** Strips the host namespace prefix for display (e.g. "Excel.Range" -> "Range"). */
function displayType(qualified: string): string {
	const dot = qualified.lastIndexOf('.');
	return dot >= 0 ? qualified.slice(dot + 1) : qualified;
}

/**
 * Resolves a hover description for the identifier at `offset`, or undefined when
 * the cursor is not on a resolvable identifier.
 */
export function resolveHover(
	source: string,
	offset: number,
	ctx: HoverContext = {},
): HoverInfo | undefined {
	const tokens = tokenize(source);
	const idx = findIdentTokenIndex(tokens, offset);
	if (idx < 0) {
		return undefined;
	}
	const token = tokens[idx];
	const span: Span = { start: token.start, end: token.end };
	const name = token.rawText;

	// Member access: `receiver.member` - describe the host or project member.
	if (idx > 0 && tokens[idx - 1].rawText === '.') {
		const member = resolveMemberCompletions(source, token.end, ctx).find(
			(mem) => mem.name.toLowerCase() === name.toLowerCase(),
		);
		if (member) {
			return buildMemberHover(member, ctx, span);
		}
		// Unknown member - do not guess.
		return undefined;
	}

	// User symbol declared in the current module (live, no save required).
	const userHover = resolveUserSymbol(source, offset, name, ctx, span);
	if (userHover) {
		return userHover;
	}

	// Host global identifier (e.g. ThisWorkbook, Application).
	const globalType = resolveHostGlobal(name, ctx.model);
	if (globalType) {
		return {
			signature: `${name} As ${displayType(globalType)}`,
			details: ['Excel host global'],
			span,
			documentation: externalDocMarkdown(ctx, name),
		};
	}

	// Worksheet/document code name (e.g. Sheet1, ThisWorkbook component).
	const codeType = ctx.codeNames?.[name.toLowerCase()];
	if (codeType) {
		const friendly = displayType(codeType);
		return {
			signature: `${name} As ${friendly}`,
			details: [`${friendly} code name`],
			span,
			documentation: externalDocMarkdown(ctx, name),
		};
	}

	// Built-in VBA runtime function/statement (MsgBox, Left, CLng, RGB, ...).
	const runtime = resolveRuntimeFunction(name);
	if (runtime) {
		return {
			signature: runtime.signature,
			details: [`VBA runtime ${runtime.kind}`],
			span,
			documentation: externalDocMarkdown(ctx, name),
		};
	}

	return undefined;
}

/** Builds a hover for a resolved host or project object member. */
function buildMemberHover(
	member: MemberCompletion,
	ctx: HoverContext,
	span: Span,
): HoverInfo {
	const ownerName = displayType(member.owner);
	const ret = member.returns ? ` As ${displayType(member.returns)}` : '';
	const call = member.kind === 'method' ? '()' : '';
	const hostType = !!getHostType(member.owner, ctx.model);
	return {
		signature: `${ownerName}.${member.name}${call}${ret}`,
		details: [
			hostType ? `Excel host ${member.kind}` : `${ownerName} ${member.kind}`,
		],
		span,
		documentation:
			member.documentation ??
			externalDocMarkdown(ctx, member.name, ownerName),
	};
}

/**
 * Renders developer-defined external documentation for `name` (optionally within
 * `qualifier`) to Markdown, or undefined when none is registered. This is the
 * override that lets a team re-describe a curated library symbol.
 */
function externalDocMarkdown(
	ctx: HoverContext,
	name: string,
	qualifier?: string,
): string | undefined {
	const doc = ctx.docRegistry?.lookup(name, qualifier);
	return hasDocContent(doc) ? renderDocMarkdown(doc) : undefined;
}

/** Finds the index of the identifier-like token whose span covers `offset`. */
function findIdentTokenIndex(tokens: VbaToken[], offset: number): number {
	let fallback = -1;
	for (let i = 0; i < tokens.length; i += 1) {
		const t = tokens[i];
		if (offset >= t.start && offset <= t.end && isIdentLike(t)) {
			// Prefer a token that strictly contains the offset over one that only
			// touches it at a boundary (cursor sitting between two tokens).
			if (offset > t.start && offset < t.end) {
				return i;
			}
			fallback = i;
		}
	}
	return fallback;
}

/** Builds a hover for a user symbol resolved from the live module graph. */
function resolveUserSymbol(
	source: string,
	offset: number,
	name: string,
	ctx: HoverContext,
	span: Span,
): HoverInfo | undefined {
	const moduleName = ctx.moduleName ?? 'Module';
	const symbol = findUserSymbol(source, offset, name, ctx, moduleName);
	if (!symbol) {
		return undefined;
	}
	const info = buildSymbolHover(symbol, moduleName, span);
	// Precedence: an inline `'''` doc-comment on the declaration wins; otherwise
	// fall back to a developer-defined external metadata entry.
	const doc: VbaDoc | undefined =
		symbol.doc ?? ctx.docRegistry?.lookup(symbol.name, moduleName);
	if (hasDocContent(doc)) {
		info.documentation = renderDocMarkdown(doc);
	}
	return info;
}

/** Finds the user symbol the cursor resolves to, or undefined. */
function findUserSymbol(
	source: string,
	offset: number,
	name: string,
	ctx: HoverContext,
	moduleName: string,
): VbaSymbol | undefined {
	let mod;
	try {
		mod = buildModuleSymbols(moduleName, ctx.moduleKind ?? 'standard', source);
	} catch {
		return undefined;
	}
	const lower = name.toLowerCase();
	const matches = (s: VbaSymbol): boolean => s.name.toLowerCase() === lower;

	// 1. Symbols in the enclosing procedure (parameters, locals, constants).
	const enclosing = mod.all.find(
		(s) => isProcedureKind(s.kind) && contains(s.fullSpan, offset),
	);
	if (enclosing?.children) {
		const local = enclosing.children.find(matches);
		if (local) {
			return local;
		}
	}

	// 2. Module-level declarations.
	const top = (mod.root.children ?? []).find(matches);
	if (top) {
		return top;
	}

	// 3. Enum members declared anywhere in the module.
	for (const child of mod.root.children ?? []) {
		if (child.kind === 'enum') {
			const member = (child.children ?? []).find(matches);
			if (member) {
				return member;
			}
		}
	}

	return undefined;
}

const PROC_KEYWORD: Record<string, string> = {
	sub: 'Sub',
	function: 'Function',
	propertyGet: 'Property Get',
	propertyLet: 'Property Let',
	propertySet: 'Property Set',
};

/** Renders the signature line and detail lines for a user symbol. */
function buildSymbolHover(
	symbol: VbaSymbol,
	moduleName: string,
	span: Span,
): HoverInfo {
	const details: string[] = [];
	let signature: string;

	if (isProcedureKind(symbol.kind)) {
		const keyword = PROC_KEYWORD[symbol.kind] ?? 'Sub';
		const params = (symbol.children ?? [])
			.filter((c) => c.kind === 'parameter')
			.map((p) => (p.asType ? `${p.name} As ${p.asType}` : p.name))
			.join(', ');
		const ret = symbol.asType ? ` As ${symbol.asType}` : '';
		signature = `${keyword} ${symbol.name}(${params})${ret}`;
		details.push(`Declared in Module: ${moduleName}`);
		details.push(`Visibility: ${symbol.visibility ?? 'Public'}`);
	} else if (symbol.kind === 'enum') {
		signature = `Enum ${symbol.name}`;
		details.push(`Declared in Module: ${moduleName}`);
		details.push(`Visibility: ${symbol.visibility ?? 'Public'}`);
	} else if (symbol.kind === 'type') {
		signature = `Type ${symbol.name}`;
		details.push(`Declared in Module: ${moduleName}`);
		details.push(`Visibility: ${symbol.visibility ?? 'Public'}`);
	} else if (symbol.kind === 'enumMember') {
		signature = symbol.name;
		details.push(`Member of Enum ${symbol.containerName ?? ''}`.trimEnd());
	} else if (symbol.kind === 'typeField') {
		const as = symbol.asType ? ` As ${symbol.asType}` : '';
		signature = `${symbol.name}${as}`;
		details.push(`Field of Type ${symbol.containerName ?? ''}`.trimEnd());
	} else {
		// Variables, parameters, and constants.
		const prefix = symbol.kind === 'constant' ? 'Const ' : '';
		const as = symbol.asType ? ` As ${symbol.asType}` : '';
		signature = `${prefix}${symbol.name}${as}`;
		if (symbol.kind === 'parameter' || symbol.kind === 'localVariable') {
			if (symbol.containerName) {
				const role = symbol.kind === 'parameter' ? 'Parameter' : 'Local';
				details.push(`${role} in ${symbol.containerName}`);
			}
		} else {
			details.push(`Declared in Module: ${moduleName}`);
			if (symbol.visibility) {
				details.push(`Visibility: ${symbol.visibility}`);
			}
		}
	}

	return { signature, details, span };
}
