import { parseModule } from '../parser/parseModule';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../parser/nodes';
import { resolveExpressionType, type ExpressionTypeContext } from '../expression/resolveExpressionType';
import { detectEol, leadingWhitespace } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult } from './refactorTypes';

/**
 * Extract Variable: a selected expression is declared and assigned above its
 * own statement, and the selection becomes the name.
 *
 *     Debug.Print Range("A1").Value * 2
 *
 *     Dim cellValue As Variant
 *     cellValue = Range("A1").Value
 *     Debug.Print cellValue * 2
 *
 * The declared type, and whether the assignment needs `Set`, come from
 * `resolveExpressionType` (issue #61) - the same answer the declare-variable
 * quick fix writes, so the two never disagree about what an expression is.
 * An expression the analyzer cannot type is refused rather than guessed at:
 * `Dim x As Variant` where the value is an object compiles and then fails at
 * run time on the missing `Set`.
 */

export interface ExtractVariableInput {
	source: string;
	/** The selected expression. */
	span: Span;
	/** Passed through to the type resolver, so a project type resolves. */
	typeContext?: ExpressionTypeContext;
	/** Overrides the generated name, when the caller has one in mind. */
	name?: string;
}

export function extractVariable(input: ExtractVariableInput): VbaRefactorResult {
	const { source, span } = input;
	if (span.end <= span.start || !source.slice(span.start, span.end).trim()) {
		return refuse('Select an expression to extract.');
	}

	const module: ModuleNode = parseModule(source);
	const procedure = procedureAt(module, span);
	if (!procedure) {
		return refuse('Extract Variable works inside a procedure.');
	}
	if (span.start < procedure.body[0]?.span.start
		|| !fullyInside(span, procedure)) {
		return refuse('The selection reaches outside the procedure body.');
	}

	const statement = statementAt(procedure.body, span);
	if (!statement) {
		// Either the span sits between statements or it straddles two of them;
		// both mean there is no single statement to lift the expression out of.
		return refuse('Select an expression inside a single statement.');
	}
	if (statement.kind === 'VariableGroup') {
		return refuse('A declaration has nothing to extract; VBA cannot initialise a Dim.');
	}

	const typed = resolveExpressionType(source, span, input.typeContext ?? {});
	if (!typed) {
		return refuse('That selection is not an expression.');
	}
	if (!typed.complete) {
		return refuse('That is part of an expression, not a whole one.');
	}

	const eol = detectEol(source);
	const lineStart = startOfLine(source, statement.span.start);
	const indent = leadingWhitespace(source.slice(lineStart, statement.span.start));
	const name = input.name ?? uniqueName(nameFor(source.slice(span.start, span.end)), procedure, module, source);
	const set = typed.isObject ? 'Set ' : '';
	const declaration =
		`${indent}Dim ${name} As ${typed.type}${eol}`
		+ `${indent}${set}${name} = ${source.slice(span.start, span.end).trim()}${eol}`;

	return refactor(
		`Extract '${name}'`,
		[
			{ span: { start: lineStart, end: lineStart }, newText: declaration },
			{ span, newText: name },
		],
		// Where the caret lands afterwards: the declared name, which is the one
		// thing here the user is likely to want to type over.
		{ start: lineStart + indent.length + 'Dim '.length, end: lineStart + indent.length + 'Dim '.length + name.length },
	);
}

function procedureAt(module: ModuleNode, span: Span): ProcedureNode | undefined {
	return module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure'
			&& span.start >= member.span.start
			&& span.end <= member.span.end,
	);
}

function fullyInside(span: Span, procedure: ProcedureNode): boolean {
	const body = procedure.body;
	if (body.length === 0) {
		return false;
	}
	return span.start >= body[0].span.start && span.end <= body[body.length - 1].span.end;
}

/** The innermost statement-like node the span sits in. */
function statementAt(body: readonly BodyNode[], span: Span): BodyNode | undefined {
	for (const node of body) {
		if (span.start < node.span.start || span.end > node.span.end) {
			continue;
		}
		const inner = childrenOf(node);
		return (inner && statementAt(inner, span)) ?? node;
	}
	return undefined;
}

/**
 * The statements a block node carries, so the walk reaches the leaves. Every
 * block keeps a flat `body` in source order (arm headers included), which is
 * exactly what a containment walk wants.
 */
function childrenOf(node: BodyNode): BodyNode[] | undefined {
	switch (node.kind) {
		case 'IfBlock':
		case 'ForBlock':
		case 'DoBlock':
		case 'WhileBlock':
		case 'WithBlock':
		case 'SelectBlock':
			return node.body;
		default:
			return undefined;
	}
}

function startOfLine(source: string, offset: number): number {
	const before = source.lastIndexOf('\n', Math.max(offset - 1, 0));
	return before === -1 ? 0 : before + 1;
}

/**
 * A name read out of the expression itself: the last identifier in it, which is
 * usually what the value IS - `Range("A1").Value` gives `value`. Falls back to
 * `value` when the expression names nothing, e.g. `2 * 3`.
 */
function nameFor(expression: string): string {
	// String literals are not names: `Range("A1")` is a range, not an a1.
	const withoutStrings = expression.replace(/"(?:[^"]|"")*"?/g, '""');
	const names = withoutStrings.match(/[\p{L}_][\p{L}\p{M}\p{N}_]*/gu) ?? [];
	const last = [...names].reverse().find((name) => !RESERVED.has(name.toLowerCase()));
	if (!last) {
		return 'value';
	}
	return last.charAt(0).toLowerCase() + last.slice(1);
}

/** Words that would make a poor variable name, or are not names at all. */
const RESERVED = new Set([
	'true', 'false', 'nothing', 'empty', 'null', 'me', 'new', 'and', 'or', 'not',
	'mod', 'like', 'is', 'then', 'else', 'to', 'step',
]);

/** The first free name in the family, so an extraction never shadows. */
function uniqueName(base: string, procedure: ProcedureNode, module: ModuleNode, source: string): string {
	const taken = new Set<string>();
	for (const param of procedure.params) { taken.add(param.name.toLowerCase()); }
	taken.add(procedure.name.toLowerCase());
	for (const member of module.members) {
		if (member.kind === 'VariableGroup') {
			for (const decl of member.declarations) { taken.add(decl.name.toLowerCase()); }
		} else if (member.kind === 'Procedure') {
			taken.add(member.name.toLowerCase());
		}
	}
	// Locals are declared inside the body, which the module walk above does not
	// reach; the procedure's own text is the cheapest complete answer.
	const body = source.slice(procedure.span.start, procedure.span.end);
	for (const match of body.matchAll(/\b(?:Dim|Static|Const|ReDim)\s+([\p{L}_][\p{L}\p{M}\p{N}_]*)/giu)) {
		taken.add(match[1].toLowerCase());
	}

	if (!taken.has(base.toLowerCase())) {
		return base;
	}
	for (let n = 2; ; n += 1) {
		const candidate = `${base}${n}`;
		if (!taken.has(candidate.toLowerCase())) {
			return candidate;
		}
	}
}
