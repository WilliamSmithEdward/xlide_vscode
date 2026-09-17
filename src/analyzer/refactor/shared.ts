import type { BodyNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../parser/nodes';
import { classifyReferenceKinds } from '../references/referenceKinds';
import { findIdentifierOccurrences, type VbaIdentifierOccurrence } from '../../vbaSourceScan';
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
	const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(0, offset));
	const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(offset));
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
	const text = source.slice(span.start, span.end);
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
