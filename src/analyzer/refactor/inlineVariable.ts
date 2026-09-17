import { parseModule } from '../parser/parseModule';
import type { ModuleNode } from '../parser/nodes';
import { procedureAtOffset } from '../parser/nodes';
import { refactor, refuse, type VbaRefactorResult, type VbaTextEdit } from './refactorTypes';
import { assignmentAt, localDeclaration, localUsesIn, nameAt } from './shared';
import { wholeLineSpan } from '../../vbaSourceScan';
import { IDENT_RE } from '../lexer/tokenHelpers';

/**
 * Inline Variable: a local is replaced by what it was assigned, and its
 * declaration and assignment go.
 *
 *     Dim limit As Long
 *     limit = 3
 *     If n > limit Then
 *
 *     If n > 3 Then
 *
 * Only an ATOMIC value - a literal or a plain name - is inlined, and this is
 * where VBA differs from every language that has this refactoring. Elsewhere a
 * compound initialiser is inlined with brackets round it to keep precedence.
 * In VBA `Foo (x)` passes x BY VALUE where `Foo x` passes it by reference, so
 * adding brackets can change what a call does to its arguments. There is no
 * safe bracket, so a compound value is refused outright rather than
 * parenthesised.
 */

export interface InlineVariableInput {
	source: string;
	/** Offset of the caret, on the local's declaration or on any use of it. */
	offset: number;
}

export function inlineVariable(input: InlineVariableInput): VbaRefactorResult {
	const { source } = input;
	const module: ModuleNode = parseModule(source);
	const procedure = procedureAtOffset(module, input.offset);
	if (!procedure) {
		return refuse('Inline Variable works on a local, inside a procedure.');
	}

	const name = nameAt(source, input.offset);
	if (!name) {
		return refuse('Put the caret on the variable to inline.');
	}

	const group = localDeclaration(procedure.body, name);
	if (!group) {
		return refuse(`'${name}' is not a local declared in this procedure.`);
	}
	const declSpan = group.span;
	if (group.isConst) {
		return refuse(`'${name}' is a Const, which is already its value everywhere.`);
	}
	if (/^static$/i.test(group.modifier)) {
		return refuse(`'${name}' is Static, so it keeps its value between calls.`);
	}
	if (group.declarations.length > 1) {
		return refuse(`'${name}' shares its declaration line. Split it first.`);
	}
	if (group.declarations[0].isArray) {
		return refuse(`'${name}' is an array, which has no single value to inline.`);
	}

	// Every occurrence inside the procedure, typed. The declaration's own name
	// is not a use; the one write is the assignment being inlined.
	const { uses, writes } = localUsesIn(source, procedure, declSpan, name);
	if (writes.length === 0) {
		return refuse(`'${name}' is never assigned, so there is no value to inline.`);
	}
	if (writes.length > 1) {
		return refuse(`'${name}' is assigned ${writes.length} times, so it has no single value.`);
	}

	const assigned = assignmentAt(source, procedure.body, writes[0].offset, name);
	if ('refusal' in assigned) {
		return refuse(assigned.refusal);
	}
	const { assignment, value } = assigned;
	if (!isAtomic(value)) {
		return refuse(
			`'${name}' is assigned an expression, not a single value. `
			+ 'Bracketing it to keep precedence would change ByRef arguments to ByVal, so this declines.',
		);
	}
	// Every remaining use has to come AFTER the assignment; a read above it
	// reads the default value, and inlining would silently change it.
	const reads = uses.filter((occ) => occ.offset !== writes[0].offset);
	const early = reads.find((occ) => occ.offset < assignment.span.start);
	if (early) {
		return refuse(`'${name}' is read before it is assigned, where its value is not '${value}'.`);
	}

	const edits: VbaTextEdit[] = [
		{ span: wholeLineSpan(source, group.span), newText: '' },
		{ span: wholeLineSpan(source, assignment.span), newText: '' },
		...reads.map((occ) => ({
			span: { start: occ.offset, end: occ.offset + name.length },
			newText: value,
		})),
	];

	return refactor(`Inline '${name}'`, edits);
}

/**
 * A literal or a plain name - the only values that mean the same thing
 * wherever they are dropped, with no brackets and no precedence to keep.
 */
function isAtomic(value: string): boolean {
	return /^-?\d+(?:\.\d+)?$/.test(value)
		|| /^&[HO][0-9A-F]+&?$/i.test(value)
		|| /^"(?:[^"]|"")*"$/.test(value)
		|| /^#[^#]*#$/.test(value)
		|| /^(?:True|False|Nothing|Empty|Null)$/i.test(value)
		|| IDENT_RE.test(value);
}
