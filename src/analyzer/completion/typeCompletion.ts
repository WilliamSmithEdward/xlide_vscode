// Type-position completion (Phase 6: IntelliSense and Completions).
//
// Offers a list of known type names when the cursor is in a declaration type
// position, i.e. immediately after `As` (or `As New`), or after an expression
// `New`. Sources of candidate types, in priority order of how the UI groups
// them:
//   1. Project-defined types passed in by the caller (class/document/UserForm
//      module names, user `Type`s, and `Enum`s).
//   2. VBA built-in data types (MS-VBAL 5.2.3.1.4 type-spec / 7.x data types).
//   3. Excel host object-model types (Workbook/Worksheet/Range/Application).
//
// This module is pure (lexer + host model only); no `vscode` dependency, so it
// is unit-tested directly. The VS Code provider supplies the project type names.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import {
	EXCEL_OBJECT_MODEL,
	type HostObjectModel,
} from '../host/excelObjectModel';
import type { VbaProjectTypeKind } from '../symbols/symbolModel';
import {
	hasDocContent,
	renderDocMarkdown,
	type VbaDoc,
} from '../docs/docModel';

/** Where a candidate type name comes from (drives the UI icon/grouping). */
export type TypeCompletionKind =
	| 'primitive'
	| 'external'
	| 'host'
	| 'module'
	| 'class'
	| 'document'
	| 'userform'
	| 'enum'
	| 'userType'
	| 'ambiguous';

/** A single suggested type name. */
export interface TypeCompletion {
	name: string;
	kind: TypeCompletionKind;
	/** Short human-readable origin, e.g. "VBA type" or "Excel type". */
	detail: string;
	/** Project module that owns the type, when known. */
	moduleName?: string;
	/** Markdown documentation from inline or external metadata, when available. */
	documentation?: string;
}

/** A project-defined type the caller knows about (class/document/UserForm, Type, or Enum). */
export interface ProjectTypeName {
	name: string;
	kind: VbaProjectTypeKind;
	moduleName?: string;
	doc?: VbaDoc;
}

export interface TypeCompletionContext {
	/** User-defined types/classes/enums visible to the current module. */
	projectTypes?: readonly ProjectTypeName[];
	/** Host object model to draw built-in object types from. */
	model?: HostObjectModel;
}

/**
 * Canonical VBA built-in data types that are valid in an `As` clause.
 *
 * Verified against MS-VBAL.pdf, v20250520, section 5.2.3.1.4 (BUILT-IN-TYPE /
 * reserved-type-name) and the fundamental data types in section 2.1. `Decimal`
 * is intentionally omitted: it is not directly declarable in VBA (only reachable
 * via `Variant`/`CDec`), so suggesting it would mislead.
 */
export const VBA_PRIMITIVE_TYPES: readonly string[] = [
	'Boolean',
	'Byte',
	'Currency',
	'Date',
	'Double',
	'Integer',
	'Long',
	'LongLong',
	'LongPtr',
	'Object',
	'Single',
	'String',
	'Variant',
];

/** External interface types available through VBA's default OLE Automation reference. */
const OLE_AUTOMATION_TYPES: readonly TypeCompletion[] = [
	{
		name: 'IUnknown',
		kind: 'external',
		detail: 'OLE Automation type',
		moduleName: 'stdole',
	},
];

/** Non-trivia, non-comment, non-newline tokens of a prefix slice. */
function meaningfulTokens(slice: string): VbaToken[] {
	const out: VbaToken[] = [];
	for (const tok of tokenize(slice)) {
		if (tok.kind === 'comment' || tok.kind === 'newline') {
			continue;
		}
		out.push(tok);
	}
	return out;
}

function isWordToken(tok: VbaToken): boolean {
	return (
		tok.kind === 'identifier' ||
		tok.kind === 'keyword' ||
		tok.kind === 'bracketedIdentifier'
	);
}

/**
 * Returns the partial type text being typed if the cursor (end of `slice`) is in
 * a type position (after `As`, `As New`, or expression `New`), or undefined
 * when it is not.
 */
type TypePositionMode = 'declaration' | 'newDeclaration' | 'newExpression';

interface TypePosition {
	prefix: string;
	mode: TypePositionMode;
	qualifier?: string;
	memberPrefix?: string;
}

function tokenName(tok: VbaToken | undefined): string | undefined {
	if (!tok) {
		return undefined;
	}
	if (tok.kind === 'identifier' || tok.kind === 'keyword') {
		return tok.rawText;
	}
	if (tok.kind === 'bracketedIdentifier') {
		return tok.rawText.slice(1, -1);
	}
	return undefined;
}

function readPartialTypeName(tokens: readonly VbaToken[]): {
	prefix: string;
	qualifier?: string;
	memberPrefix?: string;
	beforeIndex: number;
} {
	let i = tokens.length - 1;
	let prefix = '';
	let qualifier: string | undefined;
	let memberPrefix: string | undefined;

	if (tokens[i]?.rawText === '.') {
		const qualifierName = tokenName(tokens[i - 1]);
		if (qualifierName) {
			qualifier = qualifierName;
			memberPrefix = '';
			prefix = `${qualifierName}.`;
			i -= 2;
		}
		return { prefix, qualifier, memberPrefix, beforeIndex: i };
	}

	const last = tokens[i];
	if (last && isWordToken(last)) {
		const lower = last.rawText.toLowerCase();
		if (lower !== 'as' && lower !== 'new') {
			const name = tokenName(last) ?? last.rawText;
			prefix = name;
			memberPrefix = name;
			i--;
			const qualifierName = tokens[i]?.rawText === '.'
				? tokenName(tokens[i - 1])
				: undefined;
			if (qualifierName) {
				qualifier = qualifierName;
				prefix = `${qualifierName}.${name}`;
				i -= 2;
			}
		}
	}

	return { prefix, qualifier, memberPrefix, beforeIndex: i };
}

function detectTypePosition(slice: string): TypePosition | undefined {
	const tokens = meaningfulTokens(slice);
	if (tokens.length === 0) {
		return undefined;
	}
	const partial = readPartialTypeName(tokens);
	let i = partial.beforeIndex;
	if (i < 0) {
		return undefined;
	}

	// Optional `New` between `As` and the cursor, or expression-level `New`.
	if (tokens[i].kind === 'keyword' && tokens[i].rawText.toLowerCase() === 'new') {
		i--;
		if (i < 0 || tokens[i].rawText.toLowerCase() !== 'as') {
			return { ...partial, mode: 'newExpression' };
		}
		return { ...partial, mode: 'newDeclaration' };
	}
	if (i < 0) {
		return undefined;
	}

	if (tokens[i].rawText.toLowerCase() === 'as') {
		return { ...partial, mode: 'declaration' };
	}
	return undefined;
}

/** Short host type names (e.g. "Workbook") derived from the host model types. */
export function hostTypeNames(model: HostObjectModel): string[] {
	const out: string[] = [];
	for (const qualified of Object.keys(model.types)) {
		const short = qualified.split('.').pop();
		if (short) {
			out.push(short);
		}
	}
	return out;
}

const PROJECT_KIND_DETAIL: Record<VbaProjectTypeKind, string> = {
	class: 'Class',
	document: 'Document module',
	userform: 'UserForm',
	enum: 'Enum',
	userType: 'User type',
};

function projectTypeCandidates(projectTypes: readonly ProjectTypeName[]): TypeCompletion[] {
	const grouped = new Map<string, {
		name: string;
		kinds: Set<VbaProjectTypeKind>;
		count: number;
		moduleName?: string;
		doc?: VbaDoc;
	}>();
	for (const projectType of projectTypes) {
		const key = projectType.name.toLowerCase();
		const group = grouped.get(key) ?? {
			name: projectType.name,
			kinds: new Set<VbaProjectTypeKind>(),
			count: 0,
			moduleName: projectType.moduleName,
		};
		group.kinds.add(projectType.kind);
		group.count++;
		if (!group.moduleName && projectType.moduleName) {
			group.moduleName = projectType.moduleName;
		}
		if (!group.doc && hasDocContent(projectType.doc)) {
			group.doc = projectType.doc;
		}
		grouped.set(key, group);
	}
	return [...grouped.values()].map((group) => {
		if (group.count !== 1) {
			return {
				name: group.name,
				kind: 'ambiguous',
				detail: 'Ambiguous project type',
			};
		}
		const kind = [...group.kinds][0];
		return {
			name: group.name,
			kind,
			detail: PROJECT_KIND_DETAIL[kind],
			moduleName: group.moduleName,
			documentation: group.doc ? renderDocMarkdown(group.doc) : undefined,
		};
	});
}

export function typeCompletionCandidates(
	ctx: TypeCompletionContext = {},
): TypeCompletion[] {
	const model = ctx.model ?? EXCEL_OBJECT_MODEL;
	const seen = new Set<string>();
	const out: TypeCompletion[] = [];
	const add = (
		name: string,
		kind: TypeCompletionKind,
		detail: string,
		moduleName?: string,
		documentation?: string,
	): void => {
		const key = name.toLowerCase();
		if (seen.has(key)) {
			return;
		}
		seen.add(key);
		out.push({ name, kind, detail, moduleName, documentation });
	};

	// 1. Project-defined types take precedence (can shadow a built-in name).
	for (const t of projectTypeCandidates(ctx.projectTypes ?? [])) {
		add(t.name, t.kind, t.detail, t.moduleName, t.documentation);
	}
	// 2. VBA built-in data types.
	for (const name of VBA_PRIMITIVE_TYPES) {
		add(name, 'primitive', 'VBA type');
	}
	// 3. OLE Automation interface types from the default stdole reference.
	for (const t of OLE_AUTOMATION_TYPES) {
		add(t.name, t.kind, t.detail, t.moduleName, t.documentation);
	}
	// 4. Excel host object-model types.
	for (const name of hostTypeNames(model)) {
		add(name, 'host', 'Excel type');
	}

	return out;
}

function qualifiedTypeName(name: string): { qualifier: string; member: string } | undefined {
	const dot = name.indexOf('.');
	if (dot <= 0 || dot >= name.length - 1) {
		return undefined;
	}
	return {
		qualifier: name.slice(0, dot),
		member: name.slice(dot + 1),
	};
}

function projectTypeCandidatesInModule(
	moduleName: string,
	projectTypes: readonly ProjectTypeName[] | undefined,
): TypeCompletion[] {
	const lowerModule = moduleName.toLowerCase();
	return projectTypeCandidates(
		(projectTypes ?? []).filter((type) => type.moduleName?.toLowerCase() === lowerModule),
	).map((candidate) => ({ ...candidate, moduleName }));
}

function projectModuleQualifierCandidates(
	projectTypes: readonly ProjectTypeName[] | undefined,
): TypeCompletion[] {
	const seen = new Set<string>();
	const out: TypeCompletion[] = [];
	for (const type of projectTypes ?? []) {
		if (!type.moduleName) {
			continue;
		}
		const key = type.moduleName.toLowerCase();
		if (seen.has(key)) {
			continue;
		}
		seen.add(key);
		out.push({
			name: type.moduleName,
			kind: 'module',
			detail: 'Module qualifier',
			moduleName: type.moduleName,
		});
	}
	return out;
}

function externalTypeCandidatesInModule(moduleName: string): readonly TypeCompletion[] {
	if (!/^stdole$/i.test(moduleName)) {
		return [];
	}
	return OLE_AUTOMATION_TYPES;
}

function externalModuleQualifierCandidates(): TypeCompletion[] {
	return [{
		name: 'stdole',
		kind: 'module',
		detail: 'External type-library qualifier',
		moduleName: 'stdole',
	}];
}

export function isCreatableTypeCompletion(candidate: TypeCompletion): boolean {
	return candidate.kind === 'class' || candidate.kind === 'userform';
}

function isNewExpressionCompletionCandidate(candidate: TypeCompletion): boolean {
	return isCreatableTypeCompletion(candidate) ||
		candidate.kind === 'host' ||
		candidate.kind === 'external';
}

function isAllowedForTypePosition(
	candidate: TypeCompletion,
	mode: TypePositionMode,
): boolean {
	if (mode === 'declaration') {
		return true;
	}
	if (mode === 'newExpression') {
		return isNewExpressionCompletionCandidate(candidate);
	}
	return isCreatableTypeCompletion(candidate);
}

export function resolveTypeName(
	name: string,
	ctx: TypeCompletionContext = {},
): TypeCompletion | undefined {
	const qualified = qualifiedTypeName(name);
	if (qualified) {
		return [
			...projectTypeCandidatesInModule(qualified.qualifier, ctx.projectTypes),
			...externalTypeCandidatesInModule(qualified.qualifier),
		].find(
			(candidate) => candidate.name.toLowerCase() === qualified.member.toLowerCase(),
		);
	}
	const lower = name.toLowerCase();
	return typeCompletionCandidates(ctx).find((candidate) => (
		candidate.name.toLowerCase() === lower
	));
}

/**
 * Resolves the type-name completions available at `offset` in `source`. Returns
 * an empty array when the cursor is not in a declaration type position.
 *
 * Candidates are de-duplicated case-insensitively (a user type shadows a
 * built-in of the same name) and filtered by the partial text already typed.
 */
export function resolveTypeCompletions(
	source: string,
	offset: number,
	ctx: TypeCompletionContext = {},
): TypeCompletion[] {
	const pos = detectTypePosition(source.slice(0, offset));
	if (!pos) {
		return [];
	}
	const model = ctx.model ?? EXCEL_OBJECT_MODEL;
	if (pos.qualifier !== undefined) {
		const memberPrefix = (pos.memberPrefix ?? '').toLowerCase();
		return [
			...projectTypeCandidatesInModule(pos.qualifier, ctx.projectTypes),
			...externalTypeCandidatesInModule(pos.qualifier),
		]
			.filter((candidate) => isAllowedForTypePosition(candidate, pos.mode))
			.filter((candidate) => !memberPrefix || candidate.name.toLowerCase().startsWith(memberPrefix));
	}
	const prefix = pos.prefix.toLowerCase();
	const candidates = typeCompletionCandidates({ ...ctx, model })
		.filter((candidate) => isAllowedForTypePosition(candidate, pos.mode))
		.filter((candidate) => !prefix || candidate.name.toLowerCase().startsWith(prefix));
	if (pos.mode !== 'declaration') {
		return candidates;
	}
	return [
		...candidates,
		...projectModuleQualifierCandidates(ctx.projectTypes)
			.filter((candidate) => !prefix || candidate.name.toLowerCase().startsWith(prefix)),
		...externalModuleQualifierCandidates()
			.filter((candidate) => !prefix || candidate.name.toLowerCase().startsWith(prefix)),
	];
}
