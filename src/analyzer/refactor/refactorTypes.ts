import type { Span } from '../parser/nodes';
import type { VbaTextEdit } from '../codeActions/diagnosticCodeActions';

/**
 * The refactorings xlide_vbide ships beyond rename
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/69), as pure functions
 * over module source.
 *
 * Every one either produces edits or refuses with a reason. The refusals are
 * the design, not the error handling: a refactoring that proceeds where one of
 * them fires changes what the code means. VBA gives them teeth that other
 * languages do not - `Foo (x)` passes by value where `Foo x` passes by
 * reference, an undeclared name's lifetime depends on `Option Explicit`, and a
 * `Static` local outlives its call - so each refusal below names the rule it
 * protects rather than just declining.
 */

export type { VbaTextEdit };

/** Edits to a module other than the one the refactoring was invoked in. */
export interface VbaRefactorModuleEdits {
	moduleName: string;
	edits: readonly VbaTextEdit[];
}

export interface VbaRefactorSuccess {
	ok: true;
	/** What the code action offers to do, in the user's words. */
	title: string;
	edits: readonly VbaTextEdit[];
	/**
	 * Where the caret should land, when the refactoring introduces a name the
	 * user will want to type over.
	 */
	renameSpan?: Span;
	/**
	 * Call sites and moved code elsewhere in the project. A refactoring that
	 * changes a signature is only correct if these land with the edits above,
	 * so a caller applies all of them or none.
	 */
	otherModules?: readonly VbaRefactorModuleEdits[];
}

export interface VbaRefactorRefusal {
	ok: false;
	/** Why, in one sentence a reader can act on. */
	reason: string;
}

export type VbaRefactorResult = VbaRefactorSuccess | VbaRefactorRefusal;

export function refuse(reason: string): VbaRefactorRefusal {
	return { ok: false, reason };
}

export function refactor(
	title: string,
	edits: readonly VbaTextEdit[],
	renameSpan?: Span,
	otherModules?: readonly VbaRefactorModuleEdits[],
): VbaRefactorSuccess {
	return {
		ok: true,
		title,
		edits,
		...(renameSpan ? { renameSpan } : {}),
		...(otherModules && otherModules.length > 0 ? { otherModules } : {}),
	};
}

/** Applies edits to source, for tests and for callers without a text buffer. */
export function applyVbaTextEdits(source: string, edits: readonly VbaTextEdit[]): string {
	// Right to left, so an earlier edit's span still means what it meant when
	// it was computed.
	const ordered = [...edits].sort((a, b) => b.span.start - a.span.start);
	let out = source;
	for (const edit of ordered) {
		out = out.slice(0, edit.span.start) + edit.newText + out.slice(edit.span.end);
	}
	return out;
}
