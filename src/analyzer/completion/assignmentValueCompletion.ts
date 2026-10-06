import { completionLineCursorContext } from './cursorContext';
import { resolveMemberCompletionNamed, type MemberCompletionContext } from './memberAccess';
import { resolveExpressionType, resolveSourceAssignmentBindingAt } from '../expression/resolveExpressionType';
import { getHostEnumMembers, getHostType, hostDisplayName, resolveHostEnum } from '../host/hostModel';
import { resolveRuntimeFunction, resolveVbaLibraryQualifier } from '../runtime/vbaRuntime';
import type { ArgumentValueCompletion } from './argumentValueCompletion';
import type { VbaToken } from '../lexer/tokenKinds';
import { editorModuleSymbols } from '../symbols/editorModuleSymbols';
import { tokenName } from '../lexer/tokenHelpers';
import { assignmentTargetFromTokens, assignmentTargetName } from './assignmentTarget';
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
const PROPERTY_VALUE_ENUMS: Readonly<Record<string, string>> = {
	'excel.range.horizontalalignment': 'XlHAlign',
	'excel.range.verticalalignment': 'XlVAlign',
	'excel.border.linestyle': 'XlLineStyle',
	'excel.border.weight': 'XlBorderWeight',
	'excel.interior.pattern': 'XlPattern',
};

/** Assignment target when the caret follows `=` and at most a partial value. */
export function assignmentTargetAt(source: string, offset: number): VbaToken[] | undefined {
	const cursor = completionLineCursorContext(source, offset);
	if (cursor.inComment || cursor.inString) { return undefined; }
	let start = cursor.significantTokens.length;
	while (start > 0 && cursor.significantTokens[start - 1].start >= cursor.statementStart) { start--; }
	const tokens = cursor.significantTokens.slice(start);
	if (cursor.partialToken) { tokens.pop(); }
	return assignmentTargetFromTokens(tokens);
}

/** Known values accepted by an enum-valued property or variable assignment. */
export function resolveAssignmentValueCompletion(
	source: string,
	offset: number,
	ctx: AssignmentValueCompletionContext = {},
): ArgumentValueCompletion | undefined {
	const target = assignmentTargetAt(source, offset);
	if (!target?.length) { return undefined; }
	const named = assignmentTargetName(target);
	if (!named) { return undefined; }
	const last = target[named.index];
	const name = tokenName(last);
	if (!name) { return undefined; }
	const expressionCtx = {
		model: ctx.model, memberContext: ctx, moduleName: ctx.moduleName, moduleKind: ctx.moduleKind,
		projectVisibleSymbols: ctx.projectSymbols,
	};
	// Resolve the member before its index arguments, using its setter type.
	const member = target[named.index - 1]?.rawText === '.'
		? resolveMemberCompletionNamed(source, last.end, name, ctx) : undefined;
	if (member && (member.kind !== 'property' || member.access === 'read-only' || member.writable === false
		|| (!named.indexed && member.isArray))) { return undefined; }
	let sourceType: string | undefined;
	let sourceOwner: string | undefined;
	if (named.index === 0) {
		const binding = resolveSourceAssignmentBindingAt(source, { start: last.start, end: last.end }, name, expressionCtx);
		if (binding.scope === 'ambiguous') { return undefined; }
		const definitions = binding.definitions;
		const value = definitions.find(d => ['localVariable', 'moduleVariable', 'parameter'].includes(d.kind));
		const setter = definitions.find(d => d.kind === 'propertyLet');
		if (value) {
			if (named.indexed !== Boolean(value.isArray)) { return undefined; }
			sourceType = value.asType;
			sourceOwner = value.moduleName;
		} else if (setter) {
			sourceType = setter.children?.filter(d => d.kind === 'parameter').at(-1)?.asType;
			sourceOwner = setter.moduleName;
		} else if (definitions.length || named.indexed || resolveRuntimeFunction(name)) {
			return undefined; // calls, constants and getter-only properties are not writable values
		}
	}
	const type = (member && PROPERTY_VALUE_ENUMS[`${member.owner}.${member.name}`.toLowerCase()])
		?? member?.writeType ?? member?.declaredType ?? sourceType ?? resolveExpressionType(
			source, { start: target[0].start, end: last.end }, expressionCtx,
		)?.type;
	if (!type) { return undefined; }
	const ownerModule = sourceOwner ?? (member && ctx.projectClassMembers?.find(surface =>
		surface.name.toLowerCase() === member.owner.toLowerCase() || surface.moduleName.toLowerCase() === member.owner.toLowerCase())?.moduleName);
	const values = member && getHostType(member.owner, ctx.model)
		? hostEnumValues(type, ctx)
		: resolveEnumValues(source, type, ownerModule ? { ...ctx, moduleName: ownerModule } : ctx);
	return values ? { ...values, parameter: last.rawText } : undefined;
}

function hostEnumValues(type: string, ctx: AssignmentValueCompletionContext): Omit<ArgumentValueCompletion, 'parameter'> | undefined {
	if (type.toLowerCase() === 'boolean') {
		return { enumName: 'Boolean', constants: [{ name: 'True', value: -1 }, { name: 'False', value: 0 }] };
	}
	const enumeration = resolveHostEnum(type.split('.').at(-1)!, ctx.model);
	if (!enumeration) { return undefined; }
	const constants = getHostEnumMembers(enumeration.displayName, ctx.model);
	return constants.length ? { enumName: enumeration.displayName, constants, origin: 'host',
		qualifiedEnumName: `${enumeration.library ?? hostDisplayName(ctx.model)}.${enumeration.displayName}` } : undefined;
}

const LOCAL_ENUM_CACHE_MAX = 4;
const localEnumCache: { source: string; enums: Map<string, VbaSymbol> }[] = [];

function localEnums(source: string): Map<string, VbaSymbol> {
	const cached = localEnumCache.find(entry => entry.source === source);
	if (cached) { return cached.enums; }
	const enums = new Map<string, VbaSymbol>();
	// Reuse the shared editor snapshot instead of parsing a second symbol graph.
	for (const symbol of editorModuleSymbols('Module', 'standard', source).root.children ?? []) {
		if (symbol.kind === 'enum') { enums.set(symbol.name.toLowerCase(), symbol); }
	}
	localEnumCache.unshift({ source, enums });
	if (localEnumCache.length > LOCAL_ENUM_CACHE_MAX) { localEnumCache.pop(); }
	return enums;
}

/** Source/runtime/library enum lookup for assignment values. */
export function resolveEnumValues(source: string, type: string, ctx: AssignmentValueCompletionContext): Omit<ArgumentValueCompletion, 'parameter'> | undefined {
	if (type.toLowerCase() === 'boolean') { return hostEnumValues(type, ctx); }
	const bare = type.split('.').at(-1)!;
	const ownModule = type.slice(0, type.lastIndexOf('.')).toLowerCase() === (ctx.moduleName ?? 'Module').toLowerCase();
	const local = ctx.projectClassMembers === undefined && (!type.includes('.') || ownModule)
		? localEnums(source).get(bare.toLowerCase()) : undefined;
	if (local) {
		return { enumName: local.name, origin: 'source', qualifiedEnumName: `${ctx.moduleName ?? 'Module'}.${local.name}`,
			constants: (local.children ?? []).map(s => ({ name: s.name, value: s.defaultRaw, doc: s.doc })) };
	}
	const projectEnums = ctx.projectClassMembers?.filter(s => s.kind === 'enum' && s.name.toLowerCase() === bare.toLowerCase()
		&& (!type.includes('.') || type.slice(0, type.lastIndexOf('.')).toLowerCase() === s.moduleName.toLowerCase()));
	const projectEnum = projectEnums?.find(s => s.moduleName.toLowerCase() === ctx.moduleName?.toLowerCase())
		?? (projectEnums?.length === 1 ? projectEnums[0] : undefined);
	if (projectEnums && projectEnums.length > 1 && !projectEnum) { return undefined; }
	if (projectEnum) {
		return { enumName: projectEnum.name, origin: 'source', qualifiedEnumName: `${projectEnum.moduleName}.${projectEnum.name}`, constants: projectEnum.members.map(m => ({ name: m.name, doc: m.doc })) };
	}
	const qualifier = resolveVbaLibraryQualifier(type.replace(/^VBA\./i, ''));
	const runtime = qualifier?.constants?.filter(c => c.type === qualifier.name) ?? [];
	if (runtime.length) { return { enumName: runtime[0].type!, constants: runtime, origin: 'runtime', qualifiedEnumName: `VBA.${runtime[0].type!}` }; }
	const enumeration = resolveHostEnum(bare, ctx.model);
	if (type.includes('.') && type.slice(0, type.lastIndexOf('.')).toLowerCase() !==
		(enumeration?.library ?? hostDisplayName(ctx.model)).toLowerCase()) { return undefined; }
	if (!enumeration) { return undefined; }
	const constants = getHostEnumMembers(enumeration.displayName, ctx.model);
	return constants.length ? { enumName: enumeration.displayName, constants, origin: 'host',
		qualifiedEnumName: `${enumeration.library ?? hostDisplayName(ctx.model)}.${enumeration.displayName}` } : undefined;
}
