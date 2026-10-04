import { parseModule } from '../parser/parseModule';
import type { ModuleNode, ProcedureNode } from '../parser/nodes';
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
import { procedureCallBinding } from './procedureCallBinding';
import { statementRemovalSpan, mergeRemovals } from './shared';
import { firstTokenAtOrAfter, identifiersIn, tokenName } from '../lexer/tokenHelpers';
import { tokenize, tokenizeCached } from '../lexer/tokenize';
import { blockHeaderLineSpan } from '../diagnostics/walker';
import { ProjectIndex } from '../symbols/projectIndex';
import { resolveMemberDefinitionsAt, privateMemberOwnerAt, type MemberCompletionContext } from '../completion/memberAccess';
import { resolveBareIdentifierBinding } from '../symbols/nameResolution';
import type { ModuleSymbols, VbaSymbol, ModuleSymbolKind } from '../symbols/symbolModel';

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
	/** Host module roles, keyed by module name; omitted roles retain standard-module behavior. */
	moduleKinds?: Readonly<Record<string, ModuleSymbolKind>>;
}

export function introduceParameter(input: IntroduceParameterInput): VbaRefactorResult {
	const { source } = input;
	const module: ModuleNode = parseModule(source);
	const procedure = procedureAtOffset(module, input.offset);
	if (!procedure) {
		return refuse('Introduce Parameter works on a local, inside a procedure.');
	}

	if (procedure.procKind === 'PropertyLet' || procedure.procKind === 'PropertySet'
		|| procedure.procKind === 'PropertyGet' && module.members.some(member => member.kind === 'Procedure' && member !== procedure && member.name.toLowerCase() === procedure.name.toLowerCase() && (member.procKind === 'PropertyLet' || member.procKind === 'PropertySet'))) {
		return refuse('Changing this property parameter list also requires updating its assignment sites and paired accessors.');
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
	const parameter = parameterInsertion(source, procedure, name, type);
	if (!parameter) { return refuse('The procedure header has no reliable parameter-list boundary.'); }
	const edits: VbaTextEdit[] = [
		parameter,
		{ span: statementRemovalSpan(source, declaration.span), newText: '' },
		{ span: statementRemovalSpan(source, assignment.span), newText: '' },
	];

	const kinds = new Map(Object.entries(input.moduleKinds ?? {}).map(([key, kind]) => [key.toLowerCase(), kind]));
	let memberContext: MemberCompletionContext | undefined;
	let initializerCallsSelf = false;
	let unresolvedReceiver = false;
	let project: ProjectIndex | undefined;
	let symbols: ModuleSymbols | undefined;
	let owner: VbaSymbol | undefined;
	let otherModulesLoaded = false;
	const bindingProject = (includeOtherModules: boolean): ProjectIndex => {
		if (!symbols) {
			project = new ProjectIndex();
			project.setModule({moduleName: input.moduleName, moduleKind: kinds.get(input.moduleName.toLowerCase()) ?? 'standard', source});
			symbols = project.getModule(input.moduleName)!;
			owner = symbols.root.children?.find(symbol => symbol.fullSpan.start === procedure.span.start && symbol.name.toLowerCase() === procedure.name.toLowerCase());
		}
		if (includeOtherModules && !otherModulesLoaded) {
			for (const [moduleName, otherSource] of Object.entries(input.otherModuleSources ?? {})) {
				if (moduleName.toLowerCase() !== input.moduleName.toLowerCase()) { project!.setModule({moduleName, moduleKind: kinds.get(moduleName.toLowerCase()) ?? 'standard', source: otherSource}); }
			}
			otherModulesLoaded = true;
		}
		return project!;
	};
	const accepts = procedureCallBinding(source, input.moduleName, procedure, input.otherModuleSources ?? {}, () => bindingProject(true));
	const here = callSitesOf(source, procedure.name, { accept: (offset, call, qualifier, nameSpan) => {
		if (offset < procedure.span.start || offset >= procedure.span.end) { return true; }
		bindingProject(false);
		let target: boolean;
		if (qualifier) {
			if (!memberContext) {
				bindingProject(true);
				memberContext = {parsedModule: module, sourceTokens: tokenizeCached(source).filter(token => token.kind !== 'comment'), projectClassMembers: project!.projectMemberSurfaces(input.moduleName), meProjectType: symbols!.moduleKind === 'standard' ? undefined : input.moduleName, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map()};
			}
			const calledName = source.slice(nameSpan.start, nameSpan.end).replace(/^\[([^\]]+)\]$/, '$1');
			const definitions = resolveMemberDefinitionsAt(source, nameSpan.end, calledName, memberContext);
			const privateOwner = definitions.length === 0 ? privateMemberOwnerAt(source, nameSpan.end, calledName, memberContext) : undefined;
			if (definitions.length === 0 && privateOwner === undefined) { unresolvedReceiver = true; }
			target = privateOwner?.toLowerCase() === input.moduleName.toLowerCase() || definitions.length === 1 && definitions[0].moduleName.toLowerCase() === input.moduleName.toLowerCase() && definitions[0].fullSpan.start === procedure.span.start;
		} else {
			const binding = resolveBareIdentifierBinding({currentModule: symbols!, name: procedure.name, context: call ? 'call' : 'expression', enclosingProcedure: owner, offset});
			target = binding.scope !== 'ambiguous' && binding.definitions.length === 1 && binding.definitions[0] === owner;
		}
		if (target && offset >= assignment.span.start && offset < assignment.span.end) { initializerCallsSelf = true; }
		return target && !edits.some(edit => edit.newText === '' && offset >= edit.span.start && offset < edit.span.end);
	} }).filter(site => site.offset >= procedure.span.start && site.offset < procedure.span.end || accepts(input.moduleName, source, site));
	if (unresolvedReceiver) { return refuse(`A member named '${procedure.name}' inside this procedure has an unresolved receiver, so its call cannot be updated reliably.`); }
	if (initializerCallsSelf) { return refuse('The initializer calls the procedure whose signature would change, so it cannot be moved to its callers.'); }
	for (const site of here) {
		edits.push({ span: site.argumentInsert, newText: site.argumentText(value, name) });
	}

	const otherModules: VbaRefactorModuleEdits[] = [];
	for (const [otherName, otherSource] of Object.entries(input.otherModuleSources ?? {})) {
		if (otherName.toLowerCase() === input.moduleName.toLowerCase()) {
			continue;
		}
		const sites = callSitesOf(otherSource, procedure.name)
			.filter(site => accepts(otherName, otherSource, site));
		if (sites.length > 0) {
			otherModules.push({
				moduleName: otherName,
				edits: sites.map((site) => ({
					span: site.argumentInsert,
					newText: site.argumentText(value, name),
				})),
			});
		}
	}

	return refactor(
		`Introduce '${name}' as a parameter`,
		mergeRemovals(edits),
		undefined,
		otherModules,
	);
}

/** One logical-header lookup supplies both the insertion point and punctuation. */
function parameterInsertion(source: string, procedure: ProcedureNode, name: string, type: string): VbaTextEdit | undefined {
	if (!procedure.nameSpan) { return undefined; }
	const nameEnd = procedure.typeSuffixSpan?.end ?? procedure.nameSpan.end;
	const header = blockHeaderLineSpan(source, procedure.span);
	const tokens = tokenizeCached(source);
	const open = firstTokenAtOrAfter(tokens, nameEnd);
	const declared = `ByVal ${name} As ${type}`;
	if (tokens[open]?.start >= header.end || tokens[open]?.rawText !== '(') {
		return { span: { start: nameEnd, end: nameEnd }, newText: `(${declared})` };
	}
	let depth = 0;
	for (let i = open; i < tokens.length && tokens[i].start < header.end; i++) {
		if (tokens[i].rawText === '(') { depth++; }
		else if (tokens[i].rawText === ')' && --depth === 0) {
			const at = tokens[i].start;
			return { span: { start: at, end: at }, newText: (procedure.params.length === 0 ? '' : ', ') + declared };
		}
	}
	return undefined;
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
	if ((procedure.procKind === 'Function' || procedure.procKind === 'PropertyGet') && value.toLowerCase().includes(procedure.name.toLowerCase())) {
		const tokens = tokenize(value);
		if (tokens.some((token, index) => tokenName(token)?.toLowerCase() === procedure.name.toLowerCase()
			&& tokens[index - 1]?.rawText !== '.' && tokens[index + 1]?.rawText !== '(' && tokens[index + 1]?.rawText !== ':=')) {
			out.push(procedure.name);
		}
	}
	return out;
}
