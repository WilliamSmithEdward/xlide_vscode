import { parseModule } from '../parser/parseModule';
import type { BodyNode, ModuleNode, ProcedureNode, Span, VariableDeclNode, VariableGroupNode } from '../parser/nodes';
import { classifyReferenceKinds } from '../references/referenceKinds';
import { detectEol, findIdentifierOccurrences, leadingWhitespace } from '../../vbaSourceScan';
import { refactor, refuse, type VbaRefactorResult, type VbaTextEdit } from './refactorTypes';

/**
 * Extract Method: selected whole statements become a Private procedure below
 * the caller, and the selection becomes the call.
 *
 * The signature is READ, not guessed. Every local the selection touches is
 * classified from the analyzer's reference kinds (issue #55) and its position
 * relative to the selection:
 *
 * | inside the selection      | after it | becomes                     |
 * | ------------------------- | -------- | --------------------------- |
 * | read before it is written | -        | a parameter, ByVal          |
 * | read before written, and written | read | a parameter, ByRef   |
 * | written first             | read     | the result, or ByRef        |
 * | written first             | not read | its Dim moves across        |
 *
 * A name that is not a local is a free reference and needs nothing: VBA
 * resolves module and project scope from the new procedure exactly as it did
 * from the old one.
 *
 * `Option Explicit` is required. Without it an undeclared name is created on
 * first use with procedure lifetime, so moving statements into a new procedure
 * silently gives it a second, separate variable - the extraction would compile
 * and quietly do something else.
 */

export interface ExtractMethodInput {
	source: string;
	/** The selected statements. */
	span: Span;
	/** Name for the new procedure. */
	name?: string;
}

interface LocalUse {
	name: string;
	declaration?: { group: VariableGroupNode; decl: VariableDeclNode };
	isParameter: boolean;
	type: string;
	readBeforeWriteInside: boolean;
	writtenInside: boolean;
	readAfter: boolean;
	isStatic: boolean;
}

export function extractMethod(input: ExtractMethodInput): VbaRefactorResult {
	const { source } = input;
	if (input.span.end <= input.span.start || !source.slice(input.span.start, input.span.end).trim()) {
		return refuse('Select the statements to extract.');
	}
	if (!/^[ \t]*Option[ \t]+Explicit\b/im.test(source)) {
		return refuse(
			'Extract Method needs Option Explicit. Without it an undeclared name would '
			+ 'become a second, separate variable in the new procedure.',
		);
	}

	const module: ModuleNode = parseModule(source);
	const procedure = module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure'
			&& input.span.start >= member.span.start
			&& input.span.end <= member.span.end,
	);
	if (!procedure) {
		return refuse('Select statements inside one procedure.');
	}

	const selected = statementsIn(procedure.body, input.span);
	if (selected.length === 0) {
		return refuse('Select whole statements to extract.');
	}
	const block = { start: startOfLine(source, selected[0].span.start), end: selected[selected.length - 1].span.end };
	// The selection has to BE those statements, give or take whitespace: half a
	// statement cannot become a procedure body, and neither can the procedure's
	// own header or End line. Both directions matter - a selection can fall
	// short of the statements it touches, or reach past them.
	const before = input.span.start <= block.start
		? source.slice(input.span.start, block.start)
		: source.slice(block.start, input.span.start);
	const after = input.span.end >= block.end
		? source.slice(block.end, input.span.end)
		: source.slice(input.span.end, block.end);
	if (before.trim() !== '' || after.trim() !== '') {
		return /\bEnd[ \t]+(?:Sub|Function|Property)\b/i.test(after)
			|| /\b(?:Sub|Function|Property)[ \t]+[\p{L}_]/iu.test(before)
			? refuse("The selection takes in the procedure's own header or End line.")
			: refuse('Select whole statements to extract.');
	}
	if (block.start <= procedure.span.start || block.end >= endOfProcedureBody(source, procedure)) {
		return refuse("The selection takes in the procedure's own header or End line.");
	}

	const name = input.name ?? uniqueName('Extracted', module);
	if (module.members.some(
		(member) => member.kind === 'Procedure' && member.name.toLowerCase() === name.toLowerCase(),
	)) {
		return refuse(`The module already has a procedure called '${name}'.`);
	}

	const locals = classifyLocals(source, procedure, block);
	const staticLocal = locals.find((local) => local.isStatic);
	if (staticLocal) {
		return refuse(
			`'${staticLocal.name}' is Static, and a Static local keeps its value between `
			+ 'calls of the procedure it is declared in. Moving it would restart it.',
		);
	}

	const byValIn = locals.filter((l) => l.readBeforeWriteInside && !(l.writtenInside && l.readAfter));
	const byRefIn = locals.filter((l) => l.readBeforeWriteInside && l.writtenInside && l.readAfter);
	const outputs = locals.filter((l) => !l.readBeforeWriteInside && l.writtenInside && l.readAfter);
	const moved = locals.filter(
		(l) => !l.readBeforeWriteInside && l.writtenInside && !l.readAfter && l.declaration && !l.isParameter,
	);

	// One output becomes the result; more than one cannot, so they all go ByRef
	// and the extraction stays a Sub.
	const asFunction = outputs.length === 1;
	const byRefOut = asFunction ? [] : outputs;

	const params = [
		...byValIn.map((l) => ({ local: l, text: `ByVal ${l.name} As ${l.type}` })),
		...byRefIn.map((l) => ({ local: l, text: `ByRef ${l.name} As ${l.type}` })),
		...byRefOut.map((l) => ({ local: l, text: `ByRef ${l.name} As ${l.type}` })),
	];

	const eol = detectEol(source);
	const indent = leadingWhitespace(source.slice(block.start, selected[0].span.start));
	const body = source.slice(block.start, block.end);
	const movedDeclarations = moved
		.map((l) => `${indent}Dim ${l.name} As ${l.type}`)
		.join(eol);

	const header = asFunction
		? `Private Function ${name}(${params.map((p) => p.text).join(', ')}) As ${outputs[0].type}`
		: `Private Sub ${name}(${params.map((p) => p.text).join(', ')})`;
	const closer = asFunction ? 'End Function' : 'End Sub';
	// A Function returns through its own name, so the output local's last value
	// has to reach it.
	const returnLine = asFunction ? `${indent}${name} = ${outputs[0].name}` : '';

	const newProcedure = [
		header,
		...(movedDeclarations ? [movedDeclarations] : []),
		body.replace(/\s+$/, ''),
		...(returnLine ? [returnLine] : []),
		closer,
	].join(eol);

	const argumentList = params.map((p) => p.local.name).join(', ');
	const call = asFunction
		? `${indent}${outputs[0].name} = ${name}(${argumentList})`
		: `${indent}${name}${argumentList ? ` ${argumentList}` : ''}`;

	const edits: VbaTextEdit[] = [
		{ span: block, newText: call },
		// The new procedure goes below the one it came out of, which is where a
		// reader looks for a helper.
		{
			span: { start: procedure.span.end, end: procedure.span.end },
			newText: eol + eol + newProcedure + eol,
		},
	];
	// A moved Dim leaves the caller with it.
	for (const local of moved) {
		edits.push({ span: wholeLine(source, local.declaration!.group.span), newText: '' });
	}

	return refactor(`Extract '${name}'`, edits, {
		start: block.start + call.indexOf(name),
		end: block.start + call.indexOf(name) + name.length,
	});
}

/** Every local and parameter the selection touches, typed by how it is used. */
function classifyLocals(source: string, procedure: ProcedureNode, block: Span): LocalUse[] {
	const declarations = new Map<string, { group: VariableGroupNode; decl: VariableDeclNode }>();
	for (const node of walk(procedure.body)) {
		if (node.kind === 'VariableGroup') {
			for (const decl of node.declarations) {
				declarations.set(decl.name.toLowerCase(), { group: node, decl });
			}
		}
	}
	const parameters = new Map(procedure.params.map((p) => [p.name.toLowerCase(), p]));

	const out: LocalUse[] = [];
	const names = new Set([...declarations.keys(), ...parameters.keys()]);
	for (const lower of names) {
		const declaration = declarations.get(lower);
		const parameter = parameters.get(lower);
		const display = declaration?.decl.name ?? parameter?.name ?? lower;

		const occurrences = findIdentifierOccurrences(source, display)
			.filter((occ) => occ.offset >= procedure.span.start && occ.offset <= procedure.span.end)
			.filter((occ) => !declaration || !within(occ.offset, declaration.group.span));
		if (occurrences.length === 0) {
			continue;
		}
		const kinds = classifyReferenceKinds(source, occurrences.map((occ) => occ.offset));
		const inside = occurrences.filter((occ) => within(occ.offset, block));
		if (inside.length === 0) {
			continue;
		}
		out.push({
			name: display,
			...(declaration ? { declaration } : {}),
			isParameter: parameter !== undefined,
			type: declaration?.decl.asType ?? parameter?.asType ?? 'Variant',
			readBeforeWriteInside: readsBeforeAnyWrite(inside, kinds),
			writtenInside: inside.some((occ) => kinds.get(occ.offset) !== 'read'),
			readAfter: occurrences.some(
				(occ) => occ.offset > block.end && kinds.get(occ.offset) !== 'write',
			),
			isStatic: /^static$/i.test(declaration?.group.modifier ?? ''),
		});
	}
	// Source order, so a generated signature reads the way the code does.
	return out.sort((a, b) => source.indexOf(a.name) - source.indexOf(b.name));
}

/**
 * Whether the selection reads the variable before it writes it - asked per
 * STATEMENT, not per offset. `total = total + 1` writes `total` at the
 * textually first position and reads it at the second, but VBA evaluates the
 * right-hand side first, so the read comes first and the value has to arrive
 * from the caller. Ordering by offset gets that backwards and produces a
 * procedure that reads an undefined local.
 */
function readsBeforeAnyWrite(
	inside: readonly { offset: number; line: number }[],
	kinds: ReadonlyMap<number, string>,
): boolean {
	let writtenOnAnEarlierLine = false;
	for (let i = 0; i < inside.length;) {
		const line = inside[i].line;
		let reads = false;
		let writes = false;
		while (i < inside.length && inside[i].line === line) {
			const kind = kinds.get(inside[i].offset) ?? 'read';
			reads ||= kind !== 'write';
			writes ||= kind !== 'read';
			i += 1;
		}
		if (reads && !writtenOnAnEarlierLine) {
			return true;
		}
		writtenOnAnEarlierLine ||= writes;
	}
	return false;
}

/** The statements wholly or partly inside the span, at the top level of the body. */
function statementsIn(body: readonly BodyNode[], span: Span): BodyNode[] {
	const out: BodyNode[] = [];
	for (const node of body) {
		if (node.span.end < span.start || node.span.start > span.end) {
			continue;
		}
		out.push(node);
	}
	return out;
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

/** The offset just before the procedure's `End Sub` / `End Function` line. */
function endOfProcedureBody(source: string, procedure: ProcedureNode): number {
	const text = source.slice(procedure.span.start, procedure.span.end);
	const match = /[\r\n][ \t]*End[ \t]+(?:Sub|Function|Property)[ \t]*\r?\n?$/i.exec(text);
	return match ? procedure.span.start + match.index : procedure.span.end;
}

function uniqueName(base: string, module: ModuleNode): string {
	const taken = new Set(
		module.members
			.filter((member): member is ProcedureNode => member.kind === 'Procedure')
			.map((member) => member.name.toLowerCase()),
	);
	if (!taken.has(base.toLowerCase())) {
		return base;
	}
	for (let n = 2; ; n += 1) {
		if (!taken.has(`${base}${n}`.toLowerCase())) {
			return `${base}${n}`;
		}
	}
}

function within(offset: number, span: Span): boolean {
	return offset >= span.start && offset <= span.end;
}

function startOfLine(source: string, offset: number): number {
	const before = source.lastIndexOf('\n', Math.max(offset - 1, 0));
	return before === -1 ? 0 : before + 1;
}

function wholeLine(source: string, span: Span): Span {
	const start = startOfLine(source, span.start);
	const next = source.indexOf('\n', span.end);
	return { start, end: next === -1 ? source.length : next + 1 };
}
