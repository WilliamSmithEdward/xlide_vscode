// Hover resolver.
//
// Given VBA source and a cursor offset, this describes the identifier under the
// cursor: its declaration signature, kind, origin module, and visibility. It
// resolves user symbols from the live module's symbol graph and host symbols
// (globals, code names, and `receiver.member` members) from the host model.
//
// Pure analyzer code: depends only on the lexer, symbol graph, and host model,
// never on vscode. See docs/xlide_vba_language_service_roadmap.md (Phase 7).

import { tokenizeCached } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import { isIdentLike } from '../lexer/tokenHelpers';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	ModuleSymbolKind,
	VbaProcedureSignature,
	VbaSymbol,
	isProcedureKind,
	procedureDeclarationSignature,
	procedureSignatureFromSymbol,
} from '../symbols/symbolModel';
import { Span } from '../parser/nodes';
import { HostObjectModel } from '../host/excelObjectModel';
import { getHostType, resolveHostConstant, resolveHostGlobal } from '../host/hostModel';
import { resolveRuntimeConstant, resolveRuntimeFunction, resolveRuntimeObject } from '../runtime/vbaRuntime';
import {
	MemberCompletion,
	resolveMemberCompletionNamed,
} from '../completion/memberAccess';
import type { ProjectTypeName } from '../completion/typeCompletion';
import {
	resolveTypeReferenceAt,
	type ResolvedTypeReference,
} from '../semantic/typeSemanticTokens';
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
	/** Project object type that `Me` resolves to in the current class/document module. */
	meProjectType?: string;
	/** Name of the module being edited (for "Declared in Module" text). */
	moduleName?: string;
	/** Kind of the module being edited. */
	moduleKind?: ModuleSymbolKind;
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
	/** Source-declared workbook object members and visible UDT fields, keyed by type. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Project type names visible in declaration type positions. */
	projectTypes?: readonly ProjectTypeName[];
	/** Exported project procedures/Declares visible as bare calls from this module. */
	projectProcedures?: readonly VbaProcedureSignature[];
	/** Developer-defined external documentation (overrides the curated library). */
	docRegistry?: DocRegistry;
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
	const tokens = tokenizeCached(source);
	const idx = findIdentTokenIndex(tokens, offset);
	if (idx < 0) {
		return undefined;
	}
	const token = tokens[idx];
	const span: Span = { start: token.start, end: token.end };
	const name = token.rawText;

	// Type names in declarations and `New` expressions: primitives, host types,
	// and project-defined classes/enums/UDTs share the type resolver. This must
	// run before member access so qualified type names such as `Types.TPoint`
	// hover the type token rather than treating it as an expression member.
	const typeHover = resolveTypeReferenceAt(source, offset, {
		projectTypes: ctx.projectTypes,
		model: ctx.model,
	});
	if (typeHover) {
		return buildTypeHover(typeHover);
	}

	// Member access: `receiver.member` - describe the host or project member.
	if (idx > 0 && tokens[idx - 1].rawText === '.') {
		const member = resolveMemberCompletionNamed(source, token.end, name, ctx);
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

	const projectModuleHover = resolveProjectModuleHover(name, ctx, span);
	if (projectModuleHover && tokens[idx + 1]?.rawText === '.') {
		return projectModuleHover;
	}

	const projectProcedureHover = resolveProjectProcedureHover(name, ctx, span);
	if (projectProcedureHover) {
		return projectProcedureHover;
	}

	if (projectModuleHover) {
		return projectModuleHover;
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

	const runtimeObject = resolveRuntimeObject(name);
	if (runtimeObject) {
		return {
			signature: `${runtimeObject.name} As ${displayType(runtimeObject.type)}`,
			details: ['VBA runtime object'],
			span,
			documentation: externalDocMarkdown(ctx, name),
		};
	}

	const runtimeConstant = resolveRuntimeConstant(name);
	if (runtimeConstant) {
		return {
			signature: constantSignature(runtimeConstant),
			details: ['VBA runtime constant'],
			span,
			documentation: externalDocMarkdown(ctx, name),
		};
	}

	const hostConstant = resolveHostConstant(name, ctx.model);
	if (hostConstant) {
		return {
			signature: constantSignature(hostConstant),
			details: ['Excel/Office constant'],
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

function constantSignature(constant: { name: string; type?: string; value?: string | number }): string {
	const type = constant.type ? ` As ${constant.type}` : '';
	const value = constant.value !== undefined ? ` = ${formatConstantValue(constant.value)}` : '';
	return `Const ${constant.name}${type}${value}`;
}

function formatConstantValue(value: string | number): string {
	return typeof value === 'string' ? JSON.stringify(value) : String(value);
}

function buildTypeHover(typeRef: ResolvedTypeReference): HoverInfo {
	const signature = typeSignature(typeRef);
	return {
		signature,
		details: [typeDetail(typeRef)],
		span: typeRef.span,
		documentation: typeRef.documentation,
	};
}

function typeSignature(typeRef: ResolvedTypeReference): string {
	switch (typeRef.kind) {
		case 'class':
			return `Class ${typeRef.name}`;
		case 'document':
			return `Document module ${typeRef.name}`;
		case 'userform':
			return `UserForm ${typeRef.name}`;
		case 'enum':
			return `Enum ${typeRef.name}`;
		case 'userType':
			return `Type ${typeRef.name}`;
		default:
			return typeRef.name;
	}
}

function typeDetail(typeRef: ResolvedTypeReference): string {
	switch (typeRef.kind) {
		case 'primitive':
			return 'VBA primitive type';
		case 'host':
			return 'Excel host type';
		case 'external':
			return 'OLE Automation type';
		default:
			return typeRef.detail;
	}
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
	const runtimeType = !!resolveRuntimeObject(member.owner);
	const signature = member.signature
		? `${ownerName}.${member.signature}`
		: `${ownerName}.${member.name}${call}${ret}`;
	const externalDoc = externalDocMarkdown(ctx, member.name, ownerName);
	return {
		signature,
		details: [
			runtimeType
				? `VBA runtime ${member.kind}`
				: hostType ? `Excel host ${member.kind}` : `${ownerName} ${member.kind}`,
		],
		span,
		documentation: hostType
			? externalDoc ?? member.documentation
			: member.documentation ?? externalDoc,
	};
}

/**
 * Renders developer-defined external documentation for `name` (optionally within
 * `qualifier`) to Markdown, or undefined when none is registered. This is the
 * override that lets a developer re-describe a curated library symbol.
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

function resolveProjectProcedureHover(
	name: string,
	ctx: HoverContext,
	span: Span,
): HoverInfo | undefined {
	const matches = (ctx.projectProcedures ?? []).filter(
		(procedure) => procedure.name.toLowerCase() === name.toLowerCase(),
	);
	if (matches.length === 0) {
		return undefined;
	}
	if (matches.length > 1) {
		return {
			signature: name,
			details: [
				'Ambiguous project procedure',
				`Candidates: ${matches.map((procedure) => procedure.moduleName).join(', ')}`,
			],
			span,
		};
	}
	const procedure = matches[0];
	const details = [
		`Declared in Module: ${procedure.moduleName}`,
		`Visibility: ${procedure.visibility ?? 'Public'}`,
	];
	if (procedure.external) {
		details.push('External declaration');
		if (procedure.libName) {
			details.push(`Lib: ${procedure.libName}`);
		}
		if (procedure.aliasName) {
			details.push(`Alias: ${procedure.aliasName}`);
		}
	}
	const info: HoverInfo = {
		signature: procedureDeclarationSignature(procedure),
		details,
		span,
	};
	const doc: VbaDoc | undefined =
		procedure.doc ?? ctx.docRegistry?.lookup(procedure.name, procedure.moduleName);
	if (hasDocContent(doc)) {
		info.documentation = renderDocMarkdown(doc);
	}
	return info;
}

/** Signature and detail for each module surface a bare name can resolve to. */
const MODULE_SURFACE_HOVER: Record<string, { signature: (name: string) => string; detail: string }> = {
	standardModule: { signature: (name) => `Module ${name}`, detail: 'Standard module' },
	class: { signature: (name) => `Class ${name}`, detail: 'Class module' },
	document: { signature: (name) => `Document module ${name}`, detail: 'Document module' },
	userform: { signature: (name) => `UserForm ${name}`, detail: 'UserForm' },
};

function resolveProjectModuleHover(
	name: string,
	ctx: HoverContext,
	span: Span,
): HoverInfo | undefined {
	// Any object-module surface answers, not only standard modules: a class with a predeclared
	// instance (or a UserForm, or a document module) is a bare receiver too, and hovering
	// `ROneCOne` in `ROneCOne.DataView(...)` should describe it the way completion resolves it.
	// `userType` surfaces are excluded: a UDT name is not a value an expression starts from.
	const surface = (ctx.projectClassMembers ?? []).find(
		(item) => item.kind in MODULE_SURFACE_HOVER &&
			item.name.toLowerCase() === name.toLowerCase(),
	);
	if (!surface) {
		return undefined;
	}
	const hover = MODULE_SURFACE_HOVER[surface.kind];
	const doc = hasDocContent(surface.doc)
		? renderDocMarkdown(surface.doc)
		: externalDocMarkdown(ctx, surface.name);
	return {
		signature: hover.signature(surface.name),
		details: [hover.detail],
		span,
		documentation: doc,
	};
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
	} else if (symbol.kind === 'declare') {
		const callable = procedureSignatureFromSymbol(symbol);
		signature = callable ? procedureDeclarationSignature(callable) : `Declare ${symbol.name}`;
		details.push(`Declared in Module: ${moduleName}`);
		details.push(`Visibility: ${symbol.visibility ?? 'Public'}`);
		details.push('External declaration');
		if (symbol.libName) {
			details.push(`Lib: ${symbol.libName}`);
		}
		if (symbol.aliasName) {
			details.push(`Alias: ${symbol.aliasName}`);
		}
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
		const as = symbolAsClause(symbol);
		signature = `${symbol.name}${as}`;
		details.push(`Field of Type ${symbol.containerName ?? ''}`.trimEnd());
	} else {
		// Variables, parameters, and constants.
		const prefix = symbol.kind === 'constant' ? 'Const ' : '';
		const as = symbolAsClause(symbol);
		const initializer =
			symbol.kind === 'constant' && symbol.defaultRaw ? ` = ${symbol.defaultRaw}` : '';
		signature = `${prefix}${symbol.name}${as}${initializer}`;
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

function symbolAsClause(symbol: VbaSymbol): string {
	if (!symbol.asType) {
		return '';
	}
	const fixedLength = symbol.fixedLength ? ` * ${symbol.fixedLength}` : '';
	return ` As ${symbol.asType}${fixedLength}`;
}
