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
import { VbaToken } from '../lexer/tokenKinds';
import { HostObjectModel } from '../host/excelObjectModel';
import { getHostConstants, getHostGlobals, getHostType } from '../host/hostModel';
import {
	VBA_RUNTIME_CONSTANTS,
	VBA_RUNTIME_FUNCTIONS,
	type VbaRuntimeConstant,
	type VbaRuntimeFunction,
	type VbaRuntimeParam,
} from '../runtime/vbaRuntime';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	ModuleSymbolKind,
	VbaProcedureSignature,
	VbaSymbol,
	procedureDeclarationSignature,
	procedureKindKeyword,
	isProcedureKind,
} from '../symbols/symbolModel';
import { hasDocContent, renderDocMarkdown } from '../docs/docModel';

/** Origin of an identifier completion (drives the icon shown in the editor). */
export type IdentifierCompletionKind =
	| 'global'
	| 'codeName'
	| 'variable'
	| 'parameter'
	| 'constant'
	| 'procedure'
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
	/** Exported project procedures visible as bare calls from this module. */
	projectProcedures?: readonly VbaProcedureSignature[];
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

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

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
	const prefixText = source.slice(0, Math.max(0, offset));
	const tokens = tokenize(prefixText).filter((t) => t.kind !== 'comment');

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

	addInScopeSymbols(source, offset, ctx, add);

	for (const name of ctx.codeNames ?? []) {
		add(name, 'codeName', 'Worksheet object');
	}

	addProjectProcedures(ctx.projectProcedures, add);

	if (ctx.includeGlobals !== false) {
		for (const g of getHostGlobals(ctx.model)) {
			const display = getHostType(g.type, ctx.model)?.displayName ?? g.type;
			add(g.name, 'global', `${display} object`);
		}
	}

	if (ctx.includeRuntime !== false) {
		for (const f of VBA_RUNTIME_FUNCTIONS) {
			add(f.name, 'runtime', f.signature, runtimeDocumentation(f));
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
					constantDetail('Excel constant', constant.type),
					hostConstantDocumentation(constant),
				);
			}
		}
	}

	return out;
}

/**
 * Callable completions can use parentheses in expression and explicit `Call`
 * contexts, but VBA call statements like `mySub()` and `Application.Calculate()`
 * are syntax errors unless prefixed with `Call`.
 */
export function callableCompletionShouldInsertParens(
	source: string,
	offset: number,
): boolean {
	const prefixText = source.slice(0, Math.max(0, offset));
	const tokens = tokenize(prefixText).filter((t) => t.kind !== 'comment');
	if (tokens.length === 0) {
		return false;
	}

	let last = tokens.length - 1;
	if (last >= 0 && isIdentLike(tokens[last])) {
		last -= 1;
	}
	if (last < 0) {
		return false;
	}

	let boundary = last;
	while (boundary >= 0 && !isStatementBoundary(tokens[boundary])) {
		boundary -= 1;
	}
	const statement = tokens.slice(boundary + 1, last + 1);
	if (statement.length === 0) {
		return false;
	}

	const prev = statement[statement.length - 1].rawText.toLowerCase();
	if (prev === 'call' || statement[0].rawText.toLowerCase() === 'call') {
		return true;
	}
	if (statementContainsExpressionIntroducer(statement)) {
		return true;
	}
	return isExpressionContinuationToken(prev);
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
				add(member.name, 'enumMember', `${child.name} member`);
			}
		}
	}
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

function projectProcedureDocumentation(
	procedure: VbaProcedureSignature,
): string {
	const lines = [
		`**Project ${procedureKindKeyword(procedure.kind)}**`,
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
			add(symbol.name, 'procedure', 'Sub', documentation);
			return;
		case 'function':
			add(symbol.name, 'procedure', detailWithType('Function', symbol.asType), documentation);
			return;
		case 'propertyGet':
		case 'propertyLet':
		case 'propertySet':
			add(symbol.name, 'procedure', 'Property', documentation);
			return;
		case 'enum':
			add(symbol.name, 'enum', 'Enum', documentation);
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

function isStatementBoundary(token: VbaToken): boolean {
	return token.kind === 'newline' || token.rawText === ':';
}

function statementContainsExpressionIntroducer(tokens: readonly VbaToken[]): boolean {
	let depth = 0;
	for (const token of tokens) {
		if (token.rawText === '(') {
			depth += 1;
			continue;
		}
		if (token.rawText === ')') {
			depth = Math.max(0, depth - 1);
			continue;
		}
		if (depth > 0) {
			continue;
		}
		const lower = token.rawText.toLowerCase();
		if (token.rawText === '=') {
			return true;
		}
		if (
			lower === 'if' ||
			lower === 'elseif' ||
			lower === 'while' ||
			lower === 'until' ||
			lower === 'case' ||
			lower === 'select' ||
			lower === 'for' ||
			lower === 'to' ||
			lower === 'step'
		) {
			return true;
		}
	}
	return false;
}

function isExpressionContinuationToken(lowerTokenText: string): boolean {
	return (
		lowerTokenText === '(' ||
		lowerTokenText === ',' ||
		lowerTokenText === '=' ||
		lowerTokenText === '+' ||
		lowerTokenText === '-' ||
		lowerTokenText === '*' ||
		lowerTokenText === '/' ||
		lowerTokenText === '\\' ||
		lowerTokenText === '&' ||
		lowerTokenText === '<' ||
		lowerTokenText === '>' ||
		lowerTokenText === '<=' ||
		lowerTokenText === '>=' ||
		lowerTokenText === '<>' ||
		lowerTokenText === '^' ||
		lowerTokenText === 'and' ||
		lowerTokenText === 'or' ||
		lowerTokenText === 'xor' ||
		lowerTokenText === 'eqv' ||
		lowerTokenText === 'imp' ||
		lowerTokenText === 'mod' ||
		lowerTokenText === 'not' ||
		lowerTokenText === 'like' ||
		lowerTokenText === 'is'
	);
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

function hostConstantDocumentation(constant: { name: string; type?: string; value?: string | number }): string {
	const lines = [`**Excel constant**`, '', '```vba', constantSignature(constant), '```'];
	lines.push('', 'Source: generated Excel reference metadata.');
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
