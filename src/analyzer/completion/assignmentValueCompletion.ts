import { completionCursorContext } from './cursorContext';
import { resolveMemberCompletionNamed, type MemberCompletionContext } from './memberAccess';
import { resolveExpressionType } from '../expression/resolveExpressionType';
import { getHostEnumMembers, resolveHostEnum } from '../host/hostModel';
import { VBA_RUNTIME_CONSTANTS, resolveVbaLibraryQualifier } from '../runtime/vbaRuntime';
import type { ArgumentValueCompletion } from './argumentValueCompletion';
import type { VbaToken } from '../lexer/tokenKinds';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { tokenName, tokensWithoutLeadingLineNumber } from '../lexer/tokenHelpers';
import { parseExpression } from '../parser/parseExpression';
import type { VbaSymbol, ModuleSymbolKind } from '../symbols/symbolModel';

export interface AssignmentValueCompletionContext extends MemberCompletionContext {
	moduleName?: string;
	moduleKind?: ModuleSymbolKind;
	projectSymbols?: readonly VbaSymbol[];
}

// These Variant properties accept documented enums despite their typelib type.
// https://learn.microsoft.com/en-us/office/vba/api/excel.range.horizontalalignment
// https://learn.microsoft.com/en-us/office/vba/api/excel.range.verticalalignment
// https://learn.microsoft.com/en-us/office/vba/api/excel.border.linestyle
// https://learn.microsoft.com/en-us/office/vba/api/excel.border.weight
// https://learn.microsoft.com/en-us/office/vba/api/excel.interior.pattern
// https://learn.microsoft.com/en-us/office/vba/api/excel.interior.colorindex
const PROPERTY_VALUE_ENUMS: Readonly<Record<string, string>> = {
	'excel.range.horizontalalignment': 'XlHAlign',
	'excel.range.verticalalignment': 'XlVAlign',
	'excel.border.linestyle': 'XlLineStyle',
	'excel.border.weight': 'XlBorderWeight',
	'excel.interior.pattern': 'XlPattern',
	'excel.interior.colorindex': 'XlColorIndex',
};

/** Assignment target when the caret follows `=` and at most a partial value. */
export function assignmentTargetAt(source: string, offset: number): VbaToken[] | undefined {
	const cursor = completionCursorContext(source, offset);
	if (cursor.inComment || cursor.inString) { return undefined; }
	let start = cursor.significantTokens.length;
	while (start > 0 && cursor.significantTokens[start - 1].start >= cursor.statementStart) { start--; }
	const tokens = cursor.significantTokens.slice(start);
	if (cursor.partialToken) { tokens.pop(); }
	return assignmentTargetFromTokens(tokens);
}

/** Statement-local twin used by whole-document scans, avoiding prefix copies per color. */
export function assignmentTargetFromTokens(statement: readonly VbaToken[]): VbaToken[] | undefined {
	let tokens = tokensWithoutLeadingLineNumber(statement);
	if (tokens.at(-1)?.rawText !== '=') { return undefined; }
	// Only the consequent of a single-line If is an assignment; its condition is not.
	let branch = -1;
	for (let i = 0; i < tokens.length; i++) {
		if (tokens[i].kind === 'keyword' && /^(Then|Else)$/i.test(tokens[i].rawText)) { branch = i; }
	}
	if (branch >= 0) { tokens = tokens.slice(branch + 1); }
	tokens = tokens.slice(0, -1);
	if (tokens[0]?.rawText.toLowerCase() === 'let') { tokens.shift(); }
	const first = tokens[0];
	if (!first || (first.kind !== 'identifier' && first.kind !== 'bracketedIdentifier'
		&& first.rawText !== '.' && !/^(Me|ThisWorkbook)$/i.test(first.rawText))) { return undefined; }
	// The expression parser's statement callers normally bind Me as a receiver.
	// This isolated syntactic check needs only its identifier-shaped grammar.
	const syntaxTokens = tokens.map(t => t.kind === 'keyword' && t.rawText.toLowerCase() === 'me'
		? { ...t, kind: 'identifier' as const } : t);
	const parsed = parseExpression(syntaxTokens);
	return parsed.expr && parsed.endIndex === tokens.length && !parsed.diagnostics.length
		&& ['IdentifierExpr', 'MemberAccessExpr', 'IndexExpr'].includes(parsed.expr.exprKind) ? tokens : undefined;
}

export function isColorAssignmentTarget(tokens: readonly VbaToken[]): boolean {
	return tokens.at(-2)?.rawText === '.'
		&& /^(Color|BackColor|ForeColor|FillColor|BorderColor)$/i.test(tokenName(tokens.at(-1)!) ?? '');
}

/** Known values accepted by an enum-valued property or variable assignment. */
export function resolveAssignmentValueCompletion(
	source: string,
	offset: number,
	ctx: AssignmentValueCompletionContext = {},
): ArgumentValueCompletion | undefined {
	const target = assignmentTargetAt(source, offset);
	if (!target?.length) { return undefined; }
	const last = target.at(-1)!;
	// Resolve a property's assignment type from the same member surface as dot completion.
	const member = target.at(-2)?.rawText === '.'
		? resolveMemberCompletionNamed(source, last.end, tokenName(last) ?? last.rawText, ctx)
		: undefined;
	if (member?.access === 'read-only' || member?.writable === false) { return undefined; }
	const declaredType = member?.writeType ?? member?.declaredType;
	if (isColorAssignmentTarget(target) && (!declaredType || /^(Variant|Long|OLE_COLOR|stdole\.OLE_COLOR)$/i.test(declaredType))) {
		return { enumName: 'ColorConstants', parameter: last.rawText, constants: COLOR_CONSTANTS };
	}
	const type = (member && PROPERTY_VALUE_ENUMS[`${member.owner}.${member.name}`.toLowerCase()])
		?? member?.writeType ?? member?.declaredType ?? resolveExpressionType(
		source, { start: target[0].start, end: last.end }, {
			model: ctx.model, memberContext: ctx, moduleName: ctx.moduleName, moduleKind: ctx.moduleKind,
			projectVisibleSymbols: ctx.projectSymbols,
		},
	)?.type;
	if (!type) { return undefined; }
	const values = resolveEnumValues(source, type, ctx);
	return values ? { ...values, parameter: last.rawText } : undefined;
}

const COLOR_CONSTANTS = VBA_RUNTIME_CONSTANTS.filter(c => c.module === 'ColorConstants');

const LOCAL_ENUM_CACHE_MAX = 4;
const localEnumCache: { source: string; enums: Map<string, VbaSymbol> }[] = [];

function localEnums(source: string): Map<string, VbaSymbol> {
	const cached = localEnumCache.find(entry => entry.source === source);
	if (cached) { return cached.enums; }
	const enums = new Map<string, VbaSymbol>();
	// Most host/runtime enum requests need no additional source-symbol build.
	if (/\bEnum\b/i.test(source)) {
		for (const symbol of buildModuleSymbols('Module', 'standard', source).root.children ?? []) {
			if (symbol.kind === 'enum') { enums.set(symbol.name.toLowerCase(), symbol); }
		}
	}
	localEnumCache.unshift({ source, enums });
	if (localEnumCache.length > LOCAL_ENUM_CACHE_MAX) { localEnumCache.pop(); }
	return enums;
}

/** Shared source/runtime/library enum lookup for assignments and call arguments. */
export function resolveEnumValues(source: string, type: string, ctx: AssignmentValueCompletionContext): Omit<ArgumentValueCompletion, 'parameter'> | undefined {
	const bare = type.split('.').at(-1)!;
	const ownModule = type.slice(0, type.lastIndexOf('.')).toLowerCase() === (ctx.moduleName ?? 'Module').toLowerCase();
	const local = !type.includes('.') || ownModule ? localEnums(source).get(bare.toLowerCase()) : undefined;
	if (local) {
		return { enumName: local.name,
			constants: (local.children ?? []).map(s => ({ name: s.name, value: s.defaultRaw, doc: s.doc })) };
	}
	const projectEnum = ctx.projectClassMembers?.find(s => s.kind === 'enum' && s.name.toLowerCase() === bare.toLowerCase()
		&& (!type.includes('.') || type.slice(0, type.lastIndexOf('.')).toLowerCase() === s.moduleName.toLowerCase()));
	if (projectEnum) {
		return { enumName: projectEnum.name, constants: projectEnum.members.map(m => ({ name: m.name, doc: m.doc })) };
	}
	const qualifier = resolveVbaLibraryQualifier(type.replace(/^VBA\./i, ''));
	const runtime = qualifier?.constants?.filter(c => c.type === qualifier.name) ?? [];
	if (runtime.length) { return { enumName: runtime[0].type!, constants: runtime }; }
	const enumeration = resolveHostEnum(type, ctx.model);
	if (!enumeration) { return undefined; }
	const constants = getHostEnumMembers(enumeration.displayName, ctx.model);
	return constants.length ? { enumName: enumeration.displayName, constants } : undefined;
}
