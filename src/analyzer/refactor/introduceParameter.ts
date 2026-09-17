import { parseModule } from '../parser/parseModule';
import type { ModuleNode, ProcedureNode, Span } from '../parser/nodes';
import { procedureAtOffset } from '../parser/nodes';
import {
	refactor,
	refuse,
	type VbaRefactorModuleEdits,
	type VbaRefactorResult,
	type VbaTextEdit,
} from './refactorTypes';
import { assignmentAt, localDeclaration, localUsesIn, nameAt, walkBody, blankStringLiterals } from './shared';
import { callSitesOf } from './callSites';
import { wholeLineSpan } from '../../vbaSourceScan';
import { identifiersIn } from '../lexer/tokenHelpers';

/**
 * Introduce Parameter: a local becomes a `ByVal` parameter, and every call
 * site passes the value the local used to be assigned.
 *
 *     Public Sub Report()          Public Sub Report(ByVal limit As Long)
 *         Dim limit As Long            Debug.Print limit
 *         limit = 3                End Sub
 *         Debug.Print limit
 *     End Sub                      ' Report          ->  Report 3
 *
 * The refusal that matters: the initialiser has to mean the same thing at a
 * call site as it did inside the procedure. An expression naming a local, a
 * parameter, or anything Private to the module does not - it would either fail
 * to compile at the call site or, worse, bind to a different name that happens
 * to exist there. Those are refused by name rather than moved and hoped for.
 */

export interface IntroduceParameterInput {
	source: string;
	/** Offset of the caret, on the local's declaration or on a use of it. */
	offset: number;
	/** The module this source is, so call sites elsewhere can be qualified. */
	moduleName: string;
	/** Every other module in the project, keyed by name, for its call sites. */
	otherModuleSources?: Readonly<Record<string, string>>;
}

export function introduceParameter(input: IntroduceParameterInput): VbaRefactorResult {
	const { source } = input;
	const module: ModuleNode = parseModule(source);
	const procedure = procedureAtOffset(module, input.offset);
	if (!procedure) {
		return refuse('Introduce Parameter works on a local, inside a procedure.');
	}

	const name = nameAt(source, input.offset);
	if (!name) {
		return refuse('Put the caret on the local to turn into a parameter.');
	}
	if (procedure.params.some((param) => param.name.toLowerCase() === name.toLowerCase())) {
		return refuse(`'${name}' is already a parameter.`);
	}

	const declaration = localDeclaration(procedure.body, name);
	if (!declaration) {
		return refuse(`'${name}' is not a local declared in this procedure.`);
	}
	if (declaration.isConst) {
		return refuse(`'${name}' is a Const, which a caller cannot supply.`);
	}
	if (/^static$/i.test(declaration.modifier)) {
		return refuse(`'${name}' is Static, so it keeps its value between calls.`);
	}
	if (declaration.declarations.length > 1) {
		return refuse(`'${name}' shares its declaration line. Split it first.`);
	}
	const decl = declaration.declarations[0];
	if (decl.isArray) {
		return refuse(`'${name}' is an array, which cannot be passed ByVal.`);
	}

	const { writes } = localUsesIn(source, procedure, declaration.span, name);
	if (writes.length === 0) {
		return refuse(`'${name}' is never assigned, so there is no value for a caller to pass.`);
	}
	if (writes.length > 1) {
		return refuse(`'${name}' is assigned ${writes.length} times, so it has no single initial value.`);
	}

	const assigned = assignmentAt(source, procedure.body, writes[0].offset, name);
	if ('refusal' in assigned) {
		return refuse(assigned.refusal);
	}
	const { assignment, value } = assigned;
	if (source.slice(assignment.span.start, assignment.span.end).trim().match(/^Set\s/i)) {
		return refuse(`'${name}' is assigned an object with Set, which cannot be passed ByVal.`);
	}

	const stranded = strandedNames(value, procedure, module);
	if (stranded.length > 0) {
		return refuse(
			`The value '${value}' names ${stranded.map((n) => `'${n}'`).join(', ')}, `
			+ 'which a caller in another module cannot see.',
		);
	}

	const type = decl.asType ?? 'Variant';
	const edits: VbaTextEdit[] = [
		{ span: paramInsertSpan(source, procedure), newText: paramText(source, procedure, name, type) },
		{ span: wholeLineSpan(source, declaration.span), newText: '' },
		{ span: wholeLineSpan(source, assignment.span), newText: '' },
	];

	const here = callSitesOf(source, procedure.name, { skip: procedure.span });
	for (const site of here) {
		edits.push({ span: site.argumentInsert, newText: site.argumentText(value) });
	}

	const otherModules: VbaRefactorModuleEdits[] = [];
	for (const [otherName, otherSource] of Object.entries(input.otherModuleSources ?? {})) {
		if (otherName.toLowerCase() === input.moduleName.toLowerCase()) {
			continue;
		}
		const sites = callSitesOf(otherSource, procedure.name, { qualifier: input.moduleName });
		if (sites.length > 0) {
			otherModules.push({
				moduleName: otherName,
				edits: sites.map((site) => ({
					span: site.argumentInsert,
					newText: site.argumentText(value),
				})),
			});
		}
	}

	return refactor(
		`Introduce '${name}' as a parameter`,
		edits,
		undefined,
		otherModules,
	);
}

/** Where a new parameter goes, and what it looks like when others are there. */
function paramInsertSpan(source: string, procedure: ProcedureNode): Span {
	const header = source.slice(procedure.span.start, procedure.span.end);
	const close = header.indexOf(')');
	if (close === -1) {
		// `Sub Go` with no brackets at all: the whole name gets a list.
		const nameAt = header.search(new RegExp(`\\b${procedure.name}\\b`));
		const after = procedure.span.start + nameAt + procedure.name.length;
		return { start: after, end: after };
	}
	const at = procedure.span.start + close;
	return { start: at, end: at };
}

function paramText(source: string, procedure: ProcedureNode, name: string, type: string): string {
	const declared = `ByVal ${name} As ${type}`;
	const header = source.slice(procedure.span.start, procedure.span.end);
	if (header.indexOf(')') === -1) {
		return `(${declared})`;
	}
	return procedure.params.length === 0 ? declared : `, ${declared}`;
}

/**
 * Names in the initialiser that a caller elsewhere could not resolve: the
 * procedure's own locals and parameters, and anything the module keeps
 * Private. A public module member is fine - VBA resolves it project-wide.
 */
function strandedNames(value: string, procedure: ProcedureNode, module: ModuleNode): string[] {
	const locals = new Set<string>();
	for (const param of procedure.params) { locals.add(param.name.toLowerCase()); }
	for (const node of walkBody(procedure.body)) {
		if (node.kind === 'VariableGroup') {
			for (const d of node.declarations) { locals.add(d.name.toLowerCase()); }
		}
	}
	const privates = new Set<string>();
	for (const member of module.members) {
		if (member.kind === 'VariableGroup' && !/^public$/i.test(member.modifier)) {
			for (const d of member.declarations) { privates.add(d.name.toLowerCase()); }
		} else if (member.kind === 'Procedure'
			&& member.modifiers.some((m) => /^private$/i.test(m))) {
			privates.add(member.name.toLowerCase());
		}
	}

	const out: string[] = [];
	for (const name of identifiersIn(blankStringLiterals(value))) {
		const lower = name.toLowerCase();
		if ((locals.has(lower) || privates.has(lower)) && !out.includes(name)) {
			out.push(name);
		}
	}
	return out;
}
