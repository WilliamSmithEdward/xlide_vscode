import { parseModule } from '../parser/parseModule';
import type { BodyNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../parser/nodes';
import { classifyReferenceKinds } from '../references/referenceKinds';
import { findIdentifierOccurrences } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult, type VbaTextEdit } from './refactorTypes';

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
	const procedure = procedureAt(module, input.offset);
	if (!procedure) {
		return refuse('Inline Variable works on a local, inside a procedure.');
	}

	const name = nameAt(source, input.offset);
	if (!name) {
		return refuse('Put the caret on the variable to inline.');
	}

	const declaration = localDeclaration(procedure.body, name);
	if (!declaration) {
		return refuse(`'${name}' is not a local declared in this procedure.`);
	}
	const { group, span: declSpan } = declaration;
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
	const occurrences = findIdentifierOccurrences(source, name)
		.filter((occ) => occ.offset >= procedure.span.start && occ.offset <= procedure.span.end);
	const kinds = classifyReferenceKinds(source, occurrences.map((occ) => occ.offset));
	const uses = occurrences.filter((occ) => !withinSpan(occ.offset, declSpan));
	const writes = uses.filter((occ) => kinds.get(occ.offset) !== 'read');
	if (writes.length === 0) {
		return refuse(`'${name}' is never assigned, so there is no value to inline.`);
	}
	if (writes.length > 1) {
		return refuse(`'${name}' is assigned ${writes.length} times, so it has no single value.`);
	}

	const assignment = assignmentStatement(procedure.body, writes[0].offset);
	if (!assignment) {
		return refuse(`Could not find the statement that assigns '${name}'.`);
	}
	const value = assignedValue(source, assignment.span, name);
	if (value === undefined) {
		return refuse(`Could not read the value assigned to '${name}'.`);
	}
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
		{ span: wholeLine(source, group.span), newText: '' },
		{ span: wholeLine(source, assignment.span), newText: '' },
		...reads.map((occ) => ({
			span: { start: occ.offset, end: occ.offset + name.length },
			newText: value,
		})),
	];

	return refactor(`Inline '${name}'`, edits);
}

function procedureAt(module: ModuleNode, offset: number): ProcedureNode | undefined {
	return module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure' && offset >= member.span.start && offset <= member.span.end,
	);
}

/** The identifier the caret is inside, if any. */
function nameAt(source: string, offset: number): string | undefined {
	const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(0, offset));
	const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(offset));
	const name = `${before?.[0] ?? ''}${after?.[0] ?? ''}`;
	return /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(name) ? name : undefined;
}

function localDeclaration(
	body: readonly BodyNode[],
	name: string,
): { group: VariableGroupNode; span: Span } | undefined {
	const lower = name.toLowerCase();
	for (const node of walk(body)) {
		if (node.kind !== 'VariableGroup') {
			continue;
		}
		const decl = node.declarations.find((d) => d.name.toLowerCase() === lower);
		if (decl) {
			return { group: node, span: node.span };
		}
	}
	return undefined;
}

function assignmentStatement(body: readonly BodyNode[], offset: number): BodyNode | undefined {
	for (const node of walk(body)) {
		if (offset >= node.span.start && offset <= node.span.end
			&& (node.kind === 'Assignment' || node.kind === 'Statement')) {
			return node;
		}
	}
	return undefined;
}

function* walk(body: readonly BodyNode[]): Generator<BodyNode> {
	for (const node of body) {
		yield node;
		switch (node.kind) {
			case 'IfBlock':
			case 'ForBlock':
			case 'DoBlock':
			case 'WhileBlock':
			case 'WithBlock':
			case 'SelectBlock':
				yield* walk(node.body);
				break;
			default:
				break;
		}
	}
}

/** The right-hand side of `name = value` (or `Set name = value`). */
function assignedValue(source: string, span: Span, name: string): string | undefined {
	const text = source.slice(span.start, span.end);
	const match = new RegExp(
		`^\\s*(?:Set\\s+)?${escapeForRegExp(name)}\\s*=\\s*(.+?)\\s*$`,
		'i',
	).exec(text);
	return match ? match[1] : undefined;
}

function escapeForRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
		|| /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(value);
}

function withinSpan(offset: number, span: Span): boolean {
	return offset >= span.start && offset <= span.end;
}

/** The span of the whole line a statement sits on, its line break included. */
function wholeLine(source: string, span: Span): Span {
	const before = source.lastIndexOf('\n', Math.max(span.start - 1, 0));
	const start = before === -1 ? 0 : before + 1;
	const nextBreak = source.indexOf('\n', span.end);
	return { start, end: nextBreak === -1 ? source.length : nextBreak + 1 };
}
