// Type-position completion (Phase 6: IntelliSense and Completions).
//
// Offers a list of known type names when the cursor is in a declaration type
// position, i.e. immediately after `As` (or `As New`). Sources of candidate
// types, in priority order of how the UI groups them:
//   1. VBA built-in data types (MS-VBAL 5.2.3.1.4 type-spec / 7.x data types).
//   2. Excel host object-model types (Workbook/Worksheet/Range/Application).
//   3. Project-defined types passed in by the caller (user `Type`s, `Enum`s,
//      class/UserForm module names).
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

/** Where a candidate type name comes from (drives the UI icon/grouping). */
export type TypeCompletionKind =
	| 'primitive'
	| 'host'
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
}

/** A project-defined type the caller knows about (class/document/UserForm, Type, or Enum). */
export interface ProjectTypeName {
	name: string;
	kind: VbaProjectTypeKind;
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
 * a type position (after `As` or `As New`), or undefined when it is not.
 */
function detectTypePosition(slice: string): { prefix: string } | undefined {
	const tokens = meaningfulTokens(slice);
	if (tokens.length === 0) {
		return undefined;
	}
	let i = tokens.length - 1;
	let prefix = '';

	// A partial type identifier currently being typed (not `As`/`New`).
	const last = tokens[i];
	if (isWordToken(last)) {
		const lower = last.rawText.toLowerCase();
		if (lower !== 'as' && lower !== 'new') {
			prefix = last.rawText;
			i--;
		}
	}
	if (i < 0) {
		return undefined;
	}

	// Optional `New` between `As` and the cursor.
	if (tokens[i].kind === 'keyword' && tokens[i].rawText.toLowerCase() === 'new') {
		i--;
	}
	if (i < 0) {
		return undefined;
	}

	if (tokens[i].rawText.toLowerCase() === 'as') {
		return { prefix };
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
	const grouped = new Map<string, { name: string; kinds: Set<VbaProjectTypeKind> }>();
	for (const projectType of projectTypes) {
		const key = projectType.name.toLowerCase();
		const group = grouped.get(key) ?? {
			name: projectType.name,
			kinds: new Set<VbaProjectTypeKind>(),
		};
		group.kinds.add(projectType.kind);
		grouped.set(key, group);
	}
	return [...grouped.values()].map((group) => {
		if (group.kinds.size !== 1) {
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
		};
	});
}

export function typeCompletionCandidates(
	ctx: TypeCompletionContext = {},
): TypeCompletion[] {
	const model = ctx.model ?? EXCEL_OBJECT_MODEL;
	const seen = new Set<string>();
	const out: TypeCompletion[] = [];
	const add = (name: string, kind: TypeCompletionKind, detail: string): void => {
		const key = name.toLowerCase();
		if (seen.has(key)) {
			return;
		}
		seen.add(key);
		out.push({ name, kind, detail });
	};

	// 1. Project-defined types take precedence (can shadow a built-in name).
	for (const t of projectTypeCandidates(ctx.projectTypes ?? [])) {
		add(t.name, t.kind, t.detail);
	}
	// 2. VBA built-in data types.
	for (const name of VBA_PRIMITIVE_TYPES) {
		add(name, 'primitive', 'VBA type');
	}
	// 3. Excel host object-model types.
	for (const name of hostTypeNames(model)) {
		add(name, 'host', 'Excel type');
	}

	return out;
}

export function resolveTypeName(
	name: string,
	ctx: TypeCompletionContext = {},
): TypeCompletion | undefined {
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
	const prefix = pos.prefix.toLowerCase();
	return typeCompletionCandidates({ ...ctx, model }).filter((candidate) => (
		!prefix || candidate.name.toLowerCase().startsWith(prefix)
	));
}
