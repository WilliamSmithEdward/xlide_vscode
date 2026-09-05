import { parseModule } from '../parser/parseModule';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../parser/nodes';
import { detectEol, stripVba } from '../../vbaSourceScan';
import {
	refactor,
	refuse,
	type VbaRefactorModuleEdits,
	type VbaRefactorResult,
	type VbaTextEdit,
} from './refactorTypes';

/**
 * Move to Module: a procedure moves from one standard module to another.
 *
 * QUALIFIED call sites are repointed - `Reports.Build` becomes `Helpers.Build`.
 * Unqualified ones are left exactly as they are, and that is deliberate: VBA
 * resolves a bare call to a public procedure across the whole project, so
 * `Build` still reaches it from wherever it was reaching it before. Rewriting
 * those would be a large diff that changes nothing.
 *
 * The refusal names what would be stranded. A procedure that touches anything
 * Private to the module it leaves does not compile in the new one, and the
 * useful answer is which name it was - not "cannot move this".
 */

export interface MoveToModuleInput {
	source: string;
	/** Offset of the caret, inside the procedure to move. */
	offset: number;
	/** The module the procedure is leaving. */
	moduleName: string;
	/** The module it is going to. */
	targetModuleName: string;
	/** Every other module in the project, keyed by name. */
	otherModuleSources: Readonly<Record<string, string>>;
}

export function moveToModule(input: MoveToModuleInput): VbaRefactorResult {
	const { source } = input;
	const module: ModuleNode = parseModule(source);
	const procedure = module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure'
			&& input.offset >= member.span.start
			&& input.offset <= member.span.end,
	);
	if (!procedure) {
		return refuse('Put the caret in the procedure to move.');
	}
	if (input.targetModuleName.toLowerCase() === input.moduleName.toLowerCase()) {
		return refuse(`'${procedure.name}' is already in ${input.moduleName}.`);
	}

	const targetSource = lookup(input.otherModuleSources, input.targetModuleName);
	if (targetSource === undefined) {
		return refuse(`The project has no module called '${input.targetModuleName}'.`);
	}
	const target = parseModule(targetSource);
	if (target.members.some(
		(member) => member.kind === 'Procedure'
			&& member.name.toLowerCase() === procedure.name.toLowerCase(),
	)) {
		return refuse(`${input.targetModuleName} already has a procedure called '${procedure.name}'.`);
	}

	const stranded = strandedNames(source, module, procedure);
	if (stranded.length > 0) {
		return refuse(
			`'${procedure.name}' uses ${stranded.map((n) => `'${n}'`).join(', ')}, `
			+ `which ${stranded.length === 1 ? 'is' : 'are'} Private to ${input.moduleName}.`,
		);
	}

	const eol = detectEol(source);
	const moved = source.slice(procedure.span.start, procedure.span.end).replace(/\s+$/, '');
	const edits: VbaTextEdit[] = [{ span: removalSpan(source, procedure.span), newText: '' }];

	const otherModules: VbaRefactorModuleEdits[] = [{
		moduleName: input.targetModuleName,
		edits: [{
			span: { start: targetSource.length, end: targetSource.length },
			newText: (targetSource.endsWith(eol) || targetSource === '' ? '' : eol) + eol + moved + eol,
		}],
	}];

	// Qualified calls, wherever they are, including the module it leaves.
	const here = qualifiedCallEdits(source, input.moduleName, procedure.name, input.targetModuleName);
	edits.push(...here);
	for (const [name, otherSource] of Object.entries(input.otherModuleSources)) {
		if (name.toLowerCase() === input.moduleName.toLowerCase()) {
			continue;
		}
		const repoints = qualifiedCallEdits(otherSource, input.moduleName, procedure.name, input.targetModuleName);
		if (repoints.length === 0) {
			continue;
		}
		const existing = otherModules.find((m) => m.moduleName.toLowerCase() === name.toLowerCase());
		if (existing) {
			existing.edits = [...existing.edits, ...repoints];
		} else {
			otherModules.push({ moduleName: name, edits: repoints });
		}
	}

	return refactor(
		`Move '${procedure.name}' to ${input.targetModuleName}`,
		edits,
		undefined,
		otherModules,
	);
}

/**
 * Every `Module.Procedure` in the source, as an edit repointing the qualifier.
 * Only the module name is replaced, so the call keeps its own argument syntax.
 */
function qualifiedCallEdits(
	source: string,
	fromModule: string,
	procedureName: string,
	toModule: string,
): VbaTextEdit[] {
	const pattern = new RegExp(
		`\\b(${escapeForRegExp(fromModule)})\\s*\\.\\s*${escapeForRegExp(procedureName)}\\b`,
		'gi',
	);
	const out: VbaTextEdit[] = [];
	for (const match of source.matchAll(pattern)) {
		const at = match.index ?? 0;
		if (isInsideCommentOrString(source, at)) {
			continue;
		}
		out.push({ span: { start: at, end: at + match[1].length }, newText: toModule });
	}
	return out;
}

/**
 * Names the moving procedure uses that the module keeps to itself: Private
 * variables, Consts, procedures, Types, Enums and Declares. Public ones travel
 * fine - VBA resolves them project-wide.
 */
function strandedNames(source: string, module: ModuleNode, procedure: ProcedureNode): string[] {
	const privates = new Map<string, string>();
	const keep = (name: string): void => { privates.set(name.toLowerCase(), name); };
	for (const member of module.members) {
		if (member.kind === 'VariableGroup') {
			// A module-level Dim is Private, and so is anything not marked Public.
			if (!/^public$/i.test(member.modifier)) {
				for (const decl of member.declarations) { keep(decl.name); }
			}
		} else if (member.kind === 'Procedure') {
			if (member !== procedure && member.modifiers.some((m) => /^private$/i.test(m))) {
				keep(member.name);
			}
		} else if (member.kind === 'Type' || member.kind === 'Enum' || member.kind === 'Declare') {
			if (/(^|\s)Private\s/i.test(source.slice(member.span.start, member.span.end).split(/\r?\n/)[0])) {
				keep(member.name);
			}
		}
	}
	if (privates.size === 0) {
		return [];
	}

	const out: string[] = [];
	const body = source.slice(procedure.span.start, procedure.span.end);
	for (const line of body.split(/\r\n|\r|\n/)) {
		const code = stripVba(line).replace(/"(?:[^"]|"")*"?/g, '""');
		for (const match of code.matchAll(/[\p{L}_][\p{L}\p{M}\p{N}_]*/gu)) {
			const original = privates.get(match[0].toLowerCase());
			if (original && !out.includes(original)) {
				out.push(original);
			}
		}
	}
	return out;
}

/** The procedure plus the blank line under it, so the gap does not grow. */
function removalSpan(source: string, span: Span): Span {
	const before = source.lastIndexOf('\n', Math.max(span.start - 1, 0));
	const start = before === -1 ? 0 : before + 1;
	let end = span.end;
	const after = /^[ \t]*\r?\n(?:[ \t]*\r?\n)?/.exec(source.slice(end));
	if (after) {
		end += after[0].length;
	}
	return { start, end };
}

/**
 * `stripVba` blanks comments and string bodies in place, keeping every column,
 * so a character that was there and is now a space was inside one of them.
 */
function isInsideCommentOrString(source: string, offset: number): boolean {
	const lineStart = source.lastIndexOf('\n', Math.max(offset - 1, 0)) + 1;
	const found = source.indexOf('\n', offset);
	const line = source.slice(lineStart, found === -1 ? source.length : found);
	const column = offset - lineStart;
	return line[column] !== ' ' && stripVba(line)[column] === ' ';
}

function lookup(sources: Readonly<Record<string, string>>, name: string): string | undefined {
	const lower = name.toLowerCase();
	const key = Object.keys(sources).find((k) => k.toLowerCase() === lower);
	return key === undefined ? undefined : sources[key];
}

function escapeForRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
