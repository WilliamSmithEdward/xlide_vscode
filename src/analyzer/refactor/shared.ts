import { tokenize } from '../lexer/tokenize';
import type { VbaTextEdit } from './refactorTypes';
import type { BodyNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../parser/nodes';
import { classifyReferenceKinds } from '../references/referenceKinds';
import { findIdentifierOccurrences, lineStartAtAnyBreak, lineEndAtOrAfter, wholeLineSpan, type VbaIdentifierOccurrence } from '../../vbaSourceScan';
import { IDENT_RE } from '../lexer/tokenHelpers';

/**
 * What the refactorings share: where the caret is, which procedure holds it,
 * and how a module's text is searched. Each engine used to carry its own copy
 * of these, so a fix to one now reaches them all.
 */

/** The procedure whose span wholly contains `span`. */
export function procedureContainingSpan(module: ModuleNode, span: Span): ProcedureNode | undefined {
	return module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure'
			&& span.start >= member.span.start
			&& span.end <= member.span.end,
	);
}

/** The identifier the caret is inside, if any. */
export function nameAt(source: string, offset: number): string | undefined {
	// Identifiers cannot cross a physical line break. Bound both regex inputs
	// to this line rather than scanning every earlier identifier in a large
	// class whenever typing triggers code actions.
	// Normalize exactly as String.slice does, including negative offsets.
	const integerOffset = Math.trunc(offset) || 0;
	const caret = integerOffset < 0 ? Math.max(0, source.length + integerOffset) : Math.min(source.length, integerOffset);
	const start = lineStartAtAnyBreak(source, caret);
	const end = lineEndAtOrAfter(source, caret);
	const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(start, caret));
	const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(caret, end));
	const name = `${before?.[0] ?? ''}${after?.[0] ?? ''}`;
	return IDENT_RE.test(name) ? name : undefined;
}

/** The text with every string literal reduced to `""`, so nothing inside one reads as a name. */
export function blankStringLiterals(text: string): string {
	return text.replace(/"(?:[^"]|"")*"?/g, '""');
}

/** Every node of a body, nested blocks entered depth-first after their opener. */
export function* walkBody(body: readonly BodyNode[]): Generator<BodyNode> {
	for (const node of body) {
		yield node;
		switch (node.kind) {
			case 'IfBlock':
			case 'ForBlock':
			case 'DoBlock':
			case 'WhileBlock':
			case 'WithBlock':
			case 'SelectBlock':
				yield* walkBody(node.body);
				break;
			default:
				break;
		}
	}
}

/** The `Dim` group that declares `name` anywhere in the body. */
export function localDeclaration(body: readonly BodyNode[], name: string): VariableGroupNode | undefined {
	const lower = name.toLowerCase();
	for (const node of walkBody(body)) {
		if (node.kind === 'VariableGroup'
			&& node.declarations.some((d) => d.name.toLowerCase() === lower)) {
			return node;
		}
	}
	return undefined;
}

/** The assignment or plain statement whose span holds `offset`. */
export function statementAtOffset(body: readonly BodyNode[], offset: number): BodyNode | undefined {
	for (const node of walkBody(body)) {
		if (offset >= node.span.start && offset <= node.span.end
			&& (node.kind === 'Assignment' || node.kind === 'Statement')) {
			return node;
		}
	}
	return undefined;
}

/**
 * Where a local is used inside its procedure: every use, and the writes among
 * them. The declaration's own name is not a use.
 */
export function localUsesIn(
	source: string,
	procedure: ProcedureNode,
	declarationSpan: Span,
	name: string,
): { uses: VbaIdentifierOccurrence[]; writes: VbaIdentifierOccurrence[] } {
	const occurrences = findIdentifierOccurrences(source, name)
		.filter((occ) => occ.offset >= procedure.span.start && occ.offset <= procedure.span.end);
	const kinds = classifyReferenceKinds(source, occurrences.map((occ) => occ.offset));
	const uses = occurrences.filter(
		(occ) => occ.offset < declarationSpan.start || occ.offset > declarationSpan.end,
	);
	const writes = uses.filter((occ) => kinds.get(occ.offset) !== 'read');
	return { uses, writes };
}

/** The statement assigning `name` at `offset` and the value it assigns, or why that could not be read. */
export function assignmentAt(
	source: string,
	body: readonly BodyNode[],
	offset: number,
	name: string,
): { assignment: BodyNode; value: string } | { refusal: string } {
	const assignment = statementAtOffset(body, offset);
	if (!assignment) {
		return { refusal: `Could not find the statement that assigns '${name}'.` };
	}
	const value = assignedValue(source, assignment.span, name);
	if (value === undefined) {
		return { refusal: `Could not read the value assigned to '${name}'.` };
	}
	return { assignment, value };
}

/** The right-hand side of `name = value` (or `Set name = value`). */
export function assignedValue(source: string, span: Span, name: string): string | undefined {
	const code = statementCodeSpan(source, span);
	const text = source.slice(code.start, code.end);
	const match = new RegExp(`^\\s*(?:Set\\s+)?${escapeForRegExp(name)}\\s*=\\s*(.+?)\\s*$`, 'i')
		.exec(text);
	return match ? match[1] : undefined;
}

/** A module's source by name, compared the way VBA compares names. */
export function lookupModuleSource(
	sources: Readonly<Record<string, string>>,
	name: string,
): string | undefined {
	const lower = name.toLowerCase();
	const key = Object.keys(sources).find((k) => k.toLowerCase() === lower);
	return key === undefined ? undefined : sources[key];
}

export function escapeForRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Parser statement spans can include a trailing comment; it is not code. */
function statementCodeSpan(source: string, span: Span): Span {
	const text = source.slice(span.start, span.end);
	if (!text.includes("'") && !/\brem\b/i.test(text)) { return span; }
	const tokens = tokenize(text);
	const comment = tokens.findIndex((token) => token.kind === 'comment');
	return comment < 0 ? span : { start: span.start, end: span.start + (tokens[comment - 1]?.end ?? 0) };
}

/** Remove a statement without deleting its neighbors, label or trailing comment. */
export function statementRemovalSpan(source: string, span: Span): Span {
	span = statementCodeSpan(source, span);
	const start = lineStartAtAnyBreak(source, span.start);
	const end = lineEndAtOrAfter(source, span.end);
	const before = source.slice(start, span.start), after = source.slice(span.end, end);
	if (!before.trim() && !after.trim()) { return wholeLineSpan(source, span); }
	const followingColon = /^[ \t]*:/.exec(after);
	if (followingColon) { return { start: span.start, end: span.end + followingColon[0].length }; }
	const precedingColon = /:[ \t]*$/.exec(before);
	const label = /^[ \t]*(?:\d+[ \t]+)?(?:\d+|[\p{L}_][\p{L}\p{M}\p{N}_]*)[ \t]*:[ \t]*$/u.test(before);
	if (precedingColon && !label) { return { start: start + precedingColon.index, end: span.end }; }
	return { ...span };
}

/** Adjacent deleted statements can share a separator; union their removals. */
export function mergeRemovals(edits: readonly VbaTextEdit[]): VbaTextEdit[] {
	const out = edits.filter((edit) => edit.newText !== '');
	const removals = edits.filter((edit) => edit.newText === '').sort((a, b) => a.span.start - b.span.start);
	let previous: Span | undefined;
	for (const { span } of removals) {
		if (previous && span.start <= previous.end) { previous.end = Math.max(previous.end, span.end); }
		else { previous = { ...span }; out.push({ span: previous, newText: '' }); }
	}
	return out;
}
