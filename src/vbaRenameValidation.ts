// Pre-flight checks for a rename, kept free of any `vscode` import so they can
// be unit-tested directly.
//
// A rename edits several modules at once. Writing first and discovering the
// result does not compile is the worst outcome available: the project is broken
// and the operation reported success. Everything here runs before a single edit
// is produced.

import { RESERVED_IDENTIFIERS } from './analyzer/lexer/keywordTable';
import type { VbaSymbol } from './analyzer/symbols/symbolModel';

/**
 * VBA identifiers may use any locale letter, and marks continue a name (Thai
 * and Devanagari build a letter from a base plus a combining mark). The
 * ASCII-only form of this test rejected a rename to a Cyrillic or Thai name
 * that VBA accepts - the same bug class as the encoding fix in 3.1.5 and the
 * combining-mark fix in 3.2.1.
 */
export const VBA_RENAME_NAME_RE = /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u;

/** VBA caps identifiers at 255 characters. */
const MAX_IDENTIFIER_LENGTH = 255;

export interface RenameNameProblem {
	reason: 'not-an-identifier' | 'reserved' | 'too-long' | 'collision';
	message: string;
}

/** A declaration that already owns the new name in the scope being renamed. */
export interface RenameCollision {
	moduleName: string;
	/** Human-readable container, e.g. "module Helpers" or "procedure Drive". */
	container: string;
	kind: string;
}

/**
 * Shape checks on the new name alone: a legal, non-reserved VBA identifier.
 * Returns undefined when the name is usable.
 */
export function checkRenameName(newName: string): RenameNameProblem | undefined {
	if (!VBA_RENAME_NAME_RE.test(newName)) {
		return {
			reason: 'not-an-identifier',
			message: `'${newName}' is not a valid VBA identifier.`,
		};
	}
	if (newName.length > MAX_IDENTIFIER_LENGTH) {
		return {
			reason: 'too-long',
			message: `'${newName}' is longer than the ${MAX_IDENTIFIER_LENGTH} characters VBA allows for an identifier.`,
		};
	}
	if (RESERVED_IDENTIFIERS.has(newName.toLowerCase())) {
		return {
			reason: 'reserved',
			message: `'${newName}' is a reserved VBA keyword and cannot be used as a name.`,
		};
	}
	return undefined;
}

/**
 * Finds a declaration that already owns `newName` in the same scope as the
 * declaration of `oldName` - the check that stops a rename from producing two
 * `Public Sub Beta()` in one module, a project that no longer compiles.
 *
 * Scope is the declaration's SIBLINGS, not the whole module: VBA lets a local
 * shadow a module-level name, so renaming a local to a name the module also
 * declares is legal and must not be refused.
 */
export function findRenameCollision(
	root: VbaSymbol | undefined,
	moduleName: string,
	oldName: string,
	newName: string,
): RenameCollision | undefined {
	if (!root || oldName.toLowerCase() === newName.toLowerCase()) {
		return undefined;
	}
	const oldLower = oldName.toLowerCase();
	const newLower = newName.toLowerCase();

	const visit = (
		container: VbaSymbol,
		containerLabel: string,
	): RenameCollision | undefined => {
		const children = container.children ?? [];
		const declaresOld = children.some((child) => child.name.toLowerCase() === oldLower);
		if (declaresOld) {
			const clash = children.find((child) => child.name.toLowerCase() === newLower);
			if (clash) {
				return { moduleName, container: containerLabel, kind: clash.kind };
			}
		}
		for (const child of children) {
			if (!child.children?.length) {
				continue;
			}
			const hit = visit(child, `${describeKind(child.kind)} ${child.name}`);
			if (hit) {
				return hit;
			}
		}
		return undefined;
	};

	return visit(root, `module ${moduleName}`);
}

function describeKind(kind: string): string {
	switch (kind) {
		case 'sub': return 'procedure';
		case 'function': return 'procedure';
		case 'property': return 'property';
		case 'type': return 'Type';
		case 'enum': return 'Enum';
		default: return kind;
	}
}

/** Message for a collision, naming where the clash is so it can be resolved. */
export function describeRenameCollision(collision: RenameCollision, newName: string): string {
	return `Cannot rename to '${newName}': ${collision.container} in '${collision.moduleName}' `
		+ `already declares a ${describeKind(collision.kind)} with that name. `
		+ 'The rename would produce a project that does not compile.';
}
