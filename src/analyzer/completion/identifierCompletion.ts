// Identifier completion resolver.
//
// At a statement/expression position where a bare identifier is being typed
// (not after a member-access `.` and not in an `As <type>` position), this
// offers the "existing objects" available to reference here: host-injected
// globals (ThisWorkbook, ActiveSheet, Application, ...), worksheet/document
// code names, and the user's own in-scope declarations (parameters, locals,
// module-level variables/constants, procedures, enums, types).
//
// Pure analyzer code: depends only on the lexer, the symbol builder, and the
// host model. See docs/xlide_vba_language_service_roadmap.md (Phase 6).

import { tokenize } from '../lexer/tokenize';
import { completionCursorContext } from './cursorContext';
import { HostObjectModel } from '../host/excelObjectModel';
import { getHostConstants, getHostGlobals, getHostType } from '../host/hostModel';
import {
	VBA_RUNTIME_CONSTANTS,
	VBA_RUNTIME_FUNCTIONS,
	VBA_RUNTIME_OBJECTS,
	runtimeAllowsExplicitCall,
	type VbaRuntimeConstant,
	type VbaRuntimeFunction,
	type VbaRuntimeObject,
	type VbaRuntimeParam,
} from '../runtime/vbaRuntime';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	ModuleSymbolKind,
	VbaProcedureSignature,
	VbaProjectClassMembers,
	VbaSymbol,
	procedureDeclarationSignature,
	procedureSignatureFromSymbol,
	procedureKindKeyword,
	isProcedureKind,
} from '../symbols/symbolModel';
import { hasDocContent, renderDocMarkdown } from '../docs/docModel';
import {
	isExplicitCallTargetCompletionContext,
	isIdentLike,
} from '../call/callContext';

export { callableCompletionShouldInsertParens } from '../call/callContext';

/** Origin of an identifier completion (drives the icon shown in the editor). */
export type IdentifierCompletionKind =
	| 'global'
	| 'codeName'
	| 'variable'
	| 'parameter'
	| 'constant'
	| 'value'
	| 'procedure'
	| 'module'
	| 'enum'
	| 'enumMember'
	| 'type'
	| 'runtime';

/** A single identifier-completion result. */
export interface IdentifierCompletion {
	name: string;
	kind: IdentifierCompletionKind;
	detail: string;
	documentation?: string;
}

/** Project/module facts the identifier resolver needs from outside the source. */
export interface IdentifierCompletionContext {
	/** Canonical worksheet/document code names of the workbook project. */
	codeNames?: string[];
	/** Source-backed standard modules available as module-qualified receivers. */
	projectMemberSurfaces?: readonly VbaProjectClassMembers[];
	/** Exported project procedures/Declares visible as bare calls from this module. */
	projectProcedures?: readonly VbaProcedureSignature[];
	/** Source-backed project declarations visible as bare identifiers from this module. */
	projectSymbols?: readonly VbaSymbol[];
	/** Name of the module being edited (for in-scope symbol resolution). */
	moduleName?: string;
	/** Workbook-project role of the module being edited. */
	moduleKind?: ModuleSymbolKind;
	/** Include host-injected globals (default true). */
	includeGlobals?: boolean;
	/** Include built-in VBA runtime functions (default true). */
	includeRuntime?: boolean;
	/** Host object model to resolve against. Defaults to the Excel model. */
	model?: HostObjectModel;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*[$%&!#@^]?$/;

/**
 * Keywords after which the user is naming a NEW declaration rather than
 * referencing an existing object; suggesting existing identifiers there is
 * wrong (MS-VBAL 5.2/5.3 declarations).
 */
const DECLARATION_INTRODUCERS = new Set([
	'dim',
	'const',
	'redim',
	'static',
	'public',
	'private',
	'friend',
	'global',
	'withevents',
	'sub',
	'function',
	'property',
	'type',
	'enum',
	'declare',
]);

const BOOLEAN_LITERAL_COMPLETIONS: readonly IdentifierCompletion[] = [
	{
		name: 'True',
		kind: 'value',
		detail: 'Boolean literal',
		documentation: '**Boolean literal**\n\n```vba\nTrue\n```',
	},
	{
		name: 'False',
		kind: 'value',
		detail: 'Boolean literal',
		documentation: '**Boolean literal**\n\n```vba\nFalse\n```',
	},
];

/**
 * Resolves the identifier completions available at `offset`. Returns an empty
 * array when the cursor is in a member-access position (after `.`), a type
 * position (after `As`/`New`), a declaration-name position, or inside a string.
 */
export function resolveIdentifierCompletions(
	source: string,
	offset: number,
	ctx: IdentifierCompletionContext = {},
): IdentifierCompletion[] {
	const tokens = completionCursorContext(source, offset).significantTokens;

	// Identify the partial identifier being typed (if any) and the token that
	// immediately precedes it.
	let last = tokens.length - 1;
	let partial = '';
	if (last >= 0 && isIdentLike(tokens[last])) {
		partial = tokens[last].rawText;
		last -= 1;
	} else if (last >= 0 && tokens[last].kind === 'newline') {
		// Statement start: offer everything.
		last -= 1;
	} else if (last >= 0 && !isIdentLike(tokens[last])) {
		// After an operator/paren/comma/etc. - a fresh expression position.
		// Leave `last` pointing at that token for the context check below.
	}

	const before = last >= 0 ? tokens[last] : undefined;
	if (before) {
		if (before.rawText === '.') {
			return []; // member-access position
		}
		if (isIdentLike(before)) {
			const lower = before.rawText.toLowerCase();
			if (lower === 'as' || lower === 'new') {
				return []; // type position
			}
			if (DECLARATION_INTRODUCERS.has(lower)) {
				return []; // naming a new declaration
			}
		}
	}

	const lowerPartial = partial.toLowerCase();
	const explicitCallTargetContext = isExplicitCallTargetCompletionContext(tokens, last);
	const out: IdentifierCompletion[] = [];
	const seen = new Set<string>();
	const add = (
		name: string,
		kind: IdentifierCompletionKind,
		detail: string,
		documentation?: string,
	): void => {
		if (!name || !IDENT_RE.test(name)) {
			return;
		}
		const key = name.toLowerCase();
		if (seen.has(key) || !key.startsWith(lowerPartial)) {
			return;
		}
		seen.add(key);
		out.push({ name, kind, detail, documentation });
	};

	if (isBooleanLiteralCompletionContext(tokens, last)) {
		for (const literal of BOOLEAN_LITERAL_COMPLETIONS) {
			add(literal.name, literal.kind, literal.detail, literal.documentation);
		}
	}

	addInScopeSymbols(source, offset, ctx, add);

	for (const name of ctx.codeNames ?? []) {
		add(name, 'codeName', 'Worksheet object');
	}

	addProjectModules(ctx.projectMemberSurfaces, add);
	addProjectProcedures(ctx.projectProcedures, add);
	addProjectSymbols(ctx.projectSymbols, add);

	if (ctx.includeGlobals !== false) {
		for (const g of getHostGlobals(ctx.model)) {
			const display = getHostType(g.type, ctx.model)?.displayName ?? g.type;
			add(g.name, 'global', `${display} object`);
		}
	}

	if (ctx.includeRuntime !== false) {
		for (const f of VBA_RUNTIME_FUNCTIONS) {
			if (explicitCallTargetContext && !runtimeAllowsExplicitCall(f)) {
				continue;
			}
			add(f.name, 'runtime', f.signature, runtimeDocumentation(f));
		}
		for (const object of VBA_RUNTIME_OBJECTS) {
			add(object.name, 'runtime', runtimeObjectDetail(object), runtimeObjectDocumentation(object));
		}
		if (lowerPartial.length >= 2) {
			for (const constant of VBA_RUNTIME_CONSTANTS) {
				add(
					constant.name,
					'constant',
					constantDetail('VBA constant', constant.type),
					runtimeConstantDocumentation(constant),
				);
			}
			for (const constant of getHostConstants(ctx.model)) {
				add(
					constant.name,
					'constant',
					constantDetail('Excel/Office constant', constant.type),
					hostConstantDocumentation(constant),
				);
			}
		}
	}

	return out;
}

type AddFn = (
	name: string,
	kind: IdentifierCompletionKind,
	detail: string,
	documentation?: string,
) => void;

/** Adds in-scope declared symbols (params/locals of the enclosing procedure plus
 *  module-level declarations) for the module being edited. */
function addInScopeSymbols(
	source: string,
	offset: number,
	ctx: IdentifierCompletionContext,
	add: AddFn,
): void {
	let mod;
	try {
		mod = buildModuleSymbols(
			ctx.moduleName ?? 'Module',
			ctx.moduleKind ?? 'standard',
			source,
		);
	} catch {
		return;
	}

	const enclosing = mod.all.find(
		(s) =>
			isProcedureKind(s.kind) &&
			offset >= s.fullSpan.start &&
			offset <= s.fullSpan.end,
	);

	if (enclosing) {
		for (const child of enclosing.children ?? []) {
			addSymbol(child, add);
		}
		addReturnVariableSymbol(enclosing, add);
	}

	for (const child of mod.root.children ?? []) {
		addSymbol(child, add);
		// Enum members are referenceable by their bare name.
		if (child.kind === 'enum') {
			for (const member of child.children ?? []) {
				addSymbol(member, add);
			}
		}
	}
}

function isBooleanLiteralCompletionContext(
	tokens: readonly ReturnType<typeof tokenize>[number][],
	previousIndex: number,
): boolean {
	if (previousIndex < 0) {
		return false;
	}
	const before = tokens[previousIndex];
	if (!before || before.kind === 'newline' || before.kind === 'colon') {
		return false;
	}
	if (before.kind === 'operator') {
		return true;
	}
	if (before.kind === 'punctuation') {
		return before.rawText === ',' || before.rawText === '(' || before.rawText === ';';
	}
	if (isIdentLike(before)) {
		const lower = before.rawText.toLowerCase();
		return lower === 'if' ||
			lower === 'elseif' ||
			lower === 'case' ||
			lower === 'while' ||
			lower === 'until' ||
			lower === 'not' ||
			lower === 'and' ||
			lower === 'or';
	}
	return false;
}

function addProjectProcedures(
	procedures: readonly VbaProcedureSignature[] | undefined,
	add: AddFn,
): void {
	for (const procedure of procedures ?? []) {
		const detail = `${procedureDeclarationSignature(procedure)} in ${procedure.moduleName}`;
		add(procedure.name, 'procedure', detail, projectProcedureDocumentation(procedure));
	}
}

function addProjectModules(
	surfaces: readonly VbaProjectClassMembers[] | undefined,
	add: AddFn,
): void {
	for (const surface of surfaces ?? []) {
		if (surface.kind !== 'standardModule') {
			continue;
		}
		const documentation = hasDocContent(surface.doc)
			? renderDocMarkdown(surface.doc)
			: undefined;
		add(surface.name, 'module', 'Standard module', documentation);
	}
}

function addProjectSymbols(
	symbols: readonly VbaSymbol[] | undefined,
	add: AddFn,
): void {
	for (const symbol of symbols ?? []) {
		addSymbol(symbol, add);
	}
}

function projectProcedureDocumentation(
	procedure: VbaProcedureSignature,
): string {
	const kind = procedure.external
		? `Declare ${procedureKindKeyword(procedure.kind)}`
		: procedureKindKeyword(procedure.kind);
	const lines = [
		`**Project ${kind}**`,
		'',
		'```vba',
		procedureDeclarationSignature(procedure),
		'```',
		'',
		`Declared in module: \`${procedure.moduleName}\``,
	];
	const doc = hasDocContent(procedure.doc)
		? renderDocMarkdown(procedure.doc)
		: undefined;
	if (doc) {
		lines.push('', doc);
	}
	return lines.join('\n');
}

function addReturnVariableSymbol(symbol: VbaSymbol, add: AddFn): void {
	if (symbol.kind !== 'function' && symbol.kind !== 'propertyGet') {
		return;
	}
	const documentation = hasDocContent(symbol.doc)
		? renderDocMarkdown(symbol.doc)
		: undefined;
	const base = symbol.kind === 'propertyGet' ? 'Property Get return' : 'Function return';
	add(symbol.name, 'variable', detailWithType(base, symbol.asType), documentation);
}

function addSymbol(symbol: VbaSymbol, add: AddFn): void {
	const documentation = hasDocContent(symbol.doc)
		? renderDocMarkdown(symbol.doc)
		: undefined;
	switch (symbol.kind) {
		case 'parameter':
			add(symbol.name, 'parameter', detailWithType('parameter', symbol.asType), documentation);
			return;
		case 'localVariable':
			add(symbol.name, 'variable', detailWithType('local variable', symbol.asType), documentation);
			return;
		case 'moduleVariable':
			add(symbol.name, 'variable', detailWithType('module variable', symbol.asType), documentation);
			return;
		case 'constant':
			add(symbol.name, 'constant', detailWithType('constant', symbol.asType), documentation);
			return;
		case 'sub':
		case 'function':
		case 'declare': {
			const signature = procedureSignatureFromSymbol(symbol);
			const fallback = symbol.kind === 'sub'
				? 'Sub'
				: symbol.kind === 'declare'
					? 'Declare'
					: detailWithType('Function', symbol.asType);
			add(
				symbol.name,
				'procedure',
				signature ? procedureDeclarationSignature(signature) : fallback,
				documentation,
			);
			return;
		}
		case 'propertyGet':
		case 'propertyLet':
		case 'propertySet':
			add(symbol.name, 'procedure', 'Property', documentation);
			return;
		case 'enum':
			add(symbol.name, 'enum', 'Enum', documentation);
			return;
		case 'enumMember':
			add(
				symbol.name,
				'enumMember',
				symbol.containerName ? `${symbol.containerName} member` : 'Enum member',
				documentation,
			);
			return;
		case 'type':
			add(symbol.name, 'type', 'Type', documentation);
			return;
		default:
			return;
	}
}

function detailWithType(base: string, asType?: string): string {
	return asType ? `${base} As ${asType}` : base;
}

function runtimeDocumentation(fn: VbaRuntimeFunction): string {
	const lines = [
		`**VBA runtime ${fn.kind}**`,
		'',
		'```vba',
		fn.signature,
		'```',
	];
	if (fn.params?.length) {
		lines.push('', '**Parameters**');
		for (const param of fn.params) {
			lines.push(`- ${runtimeParamDocumentation(param)}`);
		}
	}
	lines.push('', 'Source: verified Microsoft VBA runtime metadata.');
	return lines.join('\n');
}

function constantDetail(base: string, type?: string): string {
	return type ? `${base} As ${type}` : base;
}

function runtimeConstantDocumentation(constant: VbaRuntimeConstant): string {
	const lines = [`**VBA runtime constant**`, '', '```vba', constantSignature(constant), '```'];
	lines.push('', 'Source: verified Microsoft VBA runtime metadata.');
	return lines.join('\n');
}

function runtimeObjectDetail(object: VbaRuntimeObject): string {
	return `VBA runtime object As ${object.type.replace(/^VBA\./, '')}`;
}

function runtimeObjectDocumentation(object: VbaRuntimeObject): string {
	const displayType = object.type.replace(/^VBA\./, '');
	const lines = [`**VBA runtime object**`, '', '```vba', `${object.name} As ${displayType}`, '```'];
	lines.push('', 'Source: verified Microsoft VBA runtime metadata.');
	return lines.join('\n');
}

function hostConstantDocumentation(constant: { name: string; type?: string; value?: string | number }): string {
	const lines = [`**Excel/Office constant**`, '', '```vba', constantSignature(constant), '```'];
	lines.push('', 'Source: generated Excel/Office reference metadata.');
	return lines.join('\n');
}

function constantSignature(constant: { name: string; type?: string; value?: string | number }): string {
	const type = constant.type ? ` As ${constant.type}` : '';
	const value = constant.value !== undefined ? ` = ${formatConstantValue(constant.value)}` : '';
	return `Const ${constant.name}${type}${value}`;
}

function formatConstantValue(value: string | number): string {
	return typeof value === 'string' ? JSON.stringify(value) : String(value);
}

function runtimeParamDocumentation(param: VbaRuntimeParam): string {
	const name = param.optional ? `[${param.name}]` : param.name;
	const suffix = param.paramArray ? ' ParamArray' : '';
	return param.type ? `\`${name}\` As \`${param.type}\`${suffix}` : `\`${name}\`${suffix}`;
}
