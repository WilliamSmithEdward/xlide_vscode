import { parseModule } from '../parser/parseModule';
import type { BodyNode, ModuleNode, ProcedureNode, Span, VariableDeclNode, VariableGroupNode } from '../parser/nodes';
import { classifyReferenceKinds } from '../references/referenceKinds';
import { detectEol, findIdentifierOccurrencesForNames, leadingWhitespace, lineStartAtAnyBreak, type VbaIdentifierOccurrence } from '../../vbaSourceScan';
import { applyVbaTextEdits, refactor, refuse, type VbaRefactorResult, type VbaTextEdit } from './refactorTypes';
import { procedureContainingSpan, walkBody, statementRemovalSpan, mergeRemovals, escapeForRegExp } from './shared';

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
	sourceOrder: number;
	declaration?: { group: VariableGroupNode; decl: VariableDeclNode };
	isParameter: boolean;
	type: string;
	readBeforeWriteInside: boolean;
	writtenInside: boolean;
	readAfter: boolean;
	usedBefore: boolean;
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
	const procedure = procedureContainingSpan(module, input.span);
	if (!procedure) {
		return refuse('Select statements inside one procedure.');
	}

	const selected = statementsIn(procedure.body, input.span);
	if (selected.length === 0) {
		return refuse('Select whole statements to extract.');
	}
	const block = { start: lineStartAtAnyBreak(source, selected[0].span.start), end: selected[selected.length - 1].span.end };
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
	if (block.start <= procedure.span.start || block.end > endOfProcedureBody(source, procedure)) {
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
	const movedDecls = new Set(moved.map((local) => local.declaration!.decl));
	const resultDecl = asFunction ? outputs[0].declaration?.decl : undefined;
	const callerDeclarations: string[] = [];
	const bodyEdits: VbaTextEdit[] = [];
	for (const group of walkBody(selected)) {
		if (group.kind !== 'VariableGroup' || !containsSpan(block, group.span)) { continue; }
		const inCaller = group.declarations.filter((decl) => !movedDecls.has(decl));
		const inHelper = group.declarations.filter((decl) => movedDecls.has(decl) || decl === resultDecl);
		if (inCaller.length > 0) {
			callerDeclarations.push(indent + declarationText(source, group, inCaller));
		}
		if (inHelper.length !== group.declarations.length) {
			const span = inHelper.length > 0 ? group.span : statementRemovalSpan(source, group.span);
			bodyEdits.push({
				span: { start: Math.max(block.start, span.start) - block.start, end: Math.min(block.end, span.end) - block.start },
				newText: inHelper.length > 0 ? declarationText(source, group, inHelper) : '',
			});
		}
	}
	const body = applyVbaTextEdits(source.slice(block.start, block.end), mergeRemovals(bodyEdits));
	const movedDeclarations = moved
		.filter((l) => !containsSpan(block, l.declaration!.group.span))
		.map((l) => `${indent}Dim ${source.slice(l.declaration!.decl.span.start, l.declaration!.decl.span.end)}`)
		.join(eol);

	const output = asFunction ? outputs[0] : undefined;
	const outputDeclaration = output && output.name.toLowerCase() !== name.toLowerCase()
		&& (!output.declaration || output.declaration.group.span.start < block.start
			|| output.declaration.group.span.end > block.end)
		? indent + 'Dim ' + (output.declaration
			? source.slice(output.declaration.decl.span.start, output.declaration.decl.span.end)
			: `${output.name} As ${output.type}`)
		: '';

	const header = asFunction
		? `Private Function ${name}(${params.map((p) => p.text).join(', ')}) As ${outputs[0].type}`
		: `Private Sub ${name}(${params.map((p) => p.text).join(', ')})`;
	const closer = asFunction ? 'End Function' : 'End Sub';
	// A Function returns through its own name, so the output local's last value
	// has to reach it.
	const returnLine = asFunction ? `${indent}${name} = ${outputs[0].name}` : '';

	const newProcedure = [
		header,
		...(outputDeclaration ? [outputDeclaration] : []),
		...(movedDeclarations ? [movedDeclarations] : []),
		body.replace(/\s+$/, ''),
		...(returnLine ? [returnLine] : []),
		closer,
	].join(eol);

	const argumentList = params.map((p) => p.local.name).join(', ');
	const invocation = asFunction
		? `${indent}${outputs[0].name} = ${name}(${argumentList})`
		: `${indent}${name}${argumentList ? ` ${argumentList}` : ''}`;

	const call = [...callerDeclarations, invocation].join(eol);
	const edits: VbaTextEdit[] = [
		{ span: block, newText: call },
		// The new procedure goes below the one it came out of, which is where a
		// reader looks for a helper.
		{
			span: { start: procedure.span.end, end: procedure.span.end },
			newText: eol + eol + newProcedure + eol,
		},
	];
	// Edit each declaration group once: siblings may remain in the caller.
	edits.push(...movedDeclarationEdits(source, moved, block));

	return refactor(`Extract '${name}'`, edits, {
		start: block.start + call.indexOf(name),
		end: block.start + call.indexOf(name) + name.length,
	});
}

/** Whether the selection includes an entire declaration statement. */
function containsSpan(outer: Span, inner: Span): boolean {
	return inner.start >= outer.start && inner.end <= outer.end;
}

function movedDeclarationEdits(source: string, moved: readonly LocalUse[], block: Span): VbaTextEdit[] {
	const groups = new Map<VariableGroupNode, Set<VariableDeclNode>>();
	for (const local of moved) {
		if (local.usedBefore) { continue; }
		const { group, decl } = local.declaration!;
		if (containsSpan(block, group.span)) { continue; }
		let declarations = groups.get(group);
		if (!declarations) { groups.set(group, declarations = new Set()); }
		declarations.add(decl);
	}
	const edits: VbaTextEdit[] = [];
	const removals: Span[] = [];
	for (const [group, declarations] of groups) {
		const remaining = group.declarations.filter((decl) => !declarations.has(decl));
		if (remaining.length === 0) {
			removals.push(statementRemovalSpan(source, group.span));
		} else {
			edits.push({ span: group.span, newText: declarationText(source, group, remaining) });
		}
	}
	return mergeRemovals([...edits, ...removals.map((span) => ({ span, newText: '' }))]);
}

function declarationText(source: string, group: VariableGroupNode, declarations: readonly VariableDeclNode[]): string {
	const first = group.declarations[0], last = group.declarations[group.declarations.length - 1];
	return source.slice(group.span.start, first.span.start)
		+ declarations.map((decl) => source.slice(decl.span.start, decl.span.end)).join(', ')
		+ source.slice(last.span.end, group.span.end);
}

/** Every local and parameter the selection touches, typed by how it is used. */
function classifyLocals(source: string, procedure: ProcedureNode, block: Span): LocalUse[] {
	const declarations = new Map<string, { group: VariableGroupNode; decl: VariableDeclNode }>();
	for (const node of walkBody(procedure.body)) {
		if (node.kind === 'VariableGroup') {
			for (const decl of node.declarations) {
				declarations.set(decl.name.toLowerCase(), { group: node, decl });
			}
		}
	}
	const parameters = new Map(procedure.params.map((p) => [p.name.toLowerCase(), p]));

	const out: LocalUse[] = [];
	const names = new Set([...declarations.keys(), ...parameters.keys()]);
	const foundByName = findIdentifierOccurrencesForNames(source, [...names]);
	const selected = new Map<string, {
		occurrences: VbaIdentifierOccurrence[];
		inside: VbaIdentifierOccurrence[];
	}>();
	for (const lower of names) {
		const declaration = declarations.get(lower);
		const occurrences = (foundByName.get(lower) ?? [])
			.filter((occ) => occ.offset >= procedure.span.start && occ.offset <= procedure.span.end)
			.filter((occ) => !declaration || !within(occ.offset, declaration.group.span));
		const inside = occurrences.filter((occ) => within(occ.offset, block));
		if (inside.length > 0) { selected.set(lower, { occurrences, inside }); }
	}
	// Reference classification itself scans the token stream: batch all touched
	// locals rather than repeating it for every declared name.
	const kinds = classifyReferenceKinds(source, [...selected.values()]
		.flatMap(({ occurrences }) => occurrences.map((occ) => occ.offset)));
	for (const [lower, { occurrences, inside }] of selected) {
		const declaration = declarations.get(lower);
		const parameter = parameters.get(lower);
		const display = declaration?.decl.name ?? parameter?.name ?? lower;

		out.push({
			name: display,
			sourceOrder: 0,
			...(declaration ? { declaration } : {}),
			isParameter: parameter !== undefined,
			type: declaration?.decl.asType ?? parameter?.asType ?? 'Variant',
			readBeforeWriteInside: readsBeforeAnyWrite(inside, kinds),
			writtenInside: inside.some((occ) => kinds.get(occ.offset) !== 'read'),
			usedBefore: occurrences.some((occ) => occ.offset < block.start),
			readAfter: occurrences.some(
				(occ) => occ.offset > block.end && kinds.get(occ.offset) !== 'write',
			),
			isStatic: /^static$/i.test(declaration?.group.modifier ?? ''),
		});
	}
	if (out.length > 1) {
		let rawOrder: Map<string, number> | undefined;
		if (out.length >= 64) {
			// Each declared spelling occurs by its name span. Unrelated suffix text
			// must not make an otherwise small ordering query choose batching.
			let searchEnd = 0;
			for (const local of out) {
				const span = local.declaration?.decl.nameSpan ?? parameters.get(local.name.toLowerCase())?.nameSpan;
				searchEnd = Math.max(searchEnd, span?.end ?? source.length);
			}
			if (out.length * searchEnd >= 16_000_000) {
				rawOrder = rawFirstOccurrences(source, out.map(local => local.name));
			}
		}
		for (const local of out) { local.sourceOrder = rawOrder?.get(local.name) ?? source.indexOf(local.name); }
	}
	// Keep the original raw first-occurrence order without rescanning while sorting.
	return out.sort((a, b) => a.sourceOrder - b.sourceOrder);
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

/** First raw UTF-16 substring positions, including overlaps and same-position prefixes. */
function rawFirstOccurrences(source: string, names: readonly string[]): Map<string, number> {
	const positions = new Map(names.map(name => [name, -1]));
	const nativeFallback = (): Map<string, number> => {
		for (const name of names) {
			if (positions.get(name) === -1) { positions.set(name, source.indexOf(name)); }
		}
		return positions;
	};
	const patterns = new Map<string, string>();
	let patternLength = 0;
	for (const name of names) {
		const pattern = escapeForRegExp(name);
		patternLength += pattern.length + 1;
		// Bound dictionary allocations and retain the native path for extreme inputs.
		if (patternLength > 65_536) { return nativeFallback(); }
		patterns.set(name, pattern);
	}
	interface PrefixNode { next: Map<string, PrefixNode>; name?: string; }
	const root: PrefixNode = { next: new Map() };
	for (const name of names) {
		let node = root;
		for (let i = 0; i < name.length; i++) {
			let next = node.next.get(name[i]);
			if (!next) { next = { next: new Map() }; node.next.set(name[i], next); }
			node = next;
		}
		node.name = name;
	}
	try {
		let matcher = new RegExp([...patterns.values()].join('|'), 'g');
		let remaining = names.length, rebuildAt = 1;
		let match: RegExpExecArray | null;
		while ((match = matcher.exec(source)) !== null) {
			let node: PrefixNode | undefined = root;
			for (let i = match.index; i < source.length; i++) {
				node = node.next.get(source[i]);
				if (!node) { break; }
				if (node.name !== undefined && positions.get(node.name) === -1) {
					positions.set(node.name, match.index); remaining--;
				}
			}
			if (remaining === 0) { break; }
			// Drop resolved short/common names without recompiling after every hit.
			if (names.length - remaining >= rebuildAt) {
				matcher = new RegExp(names.filter(name => positions.get(name) === -1).map(name => patterns.get(name)!).join('|'), 'g');
				while (rebuildAt <= names.length - remaining) { rebuildAt *= 2; }
			}
			// One code unit preserves matches that begin inside a longer match.
			matcher.lastIndex = match.index + 1;
		}
	} catch (error) {
		if (!(error instanceof SyntaxError || error instanceof RangeError)) { throw error; }
		return nativeFallback(); // Preserve the result if an engine rejects the dictionary.
	}
	return positions;
}
