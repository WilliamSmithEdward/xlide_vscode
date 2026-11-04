// Rule: a name reached through another VBA project the project references
// must be one that project offers (issue #1503). Measured through the VBE
// oracle in Excel, Word, PowerPoint and Access 16.0, with a referenced
// project ProbeLibrary whose standard module Service has a Public Function
// Value: `ProbeLibrary.Nope`, `ProbeLibrary.Service.Nope()` and, with Service
// marked `Option Private Module`, `ProbeLibrary.Service.Value()` and
// `ProbeLibrary.Value()` are each "Method or data member not found". So is
// `Service.Nope()` while Service is the referenced project's.
//
// Only referenced projects XLIDE read are judged (see
// symbols/referencedProjects.ts), and a name the project itself declares or
// sees comes first: a local, module or project name Service is not the
// referenced module.

import type { HostObjectModel } from '../../host/excelObjectModel';
import { resolveHostConstant, resolveHostGlobal, resolveHostGlobalMember } from '../../host/hostModel';
import { resolveRuntimeConstant, resolveRuntimeFunction, resolveRuntimeObject, resolveVbaLibraryQualifier } from '../../runtime/vbaRuntime';
import { collectTypeNameReferences } from '../../semantic/typeSemanticTokens';
import { referencedProjectNames, referencedProjectReached, type ReferencedLookupOrder, type ReferencedProjectModule, type ReferencedProjectSurface } from '../../symbols/referencedProjects';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import { absoluteSpan, forEachVariableGroup, statementTokens, tokenName, type ProcedureStatementVisitor } from '../walker';

/** The referenced projects and the standard modules they show, lowercased, for one module's analysis. */
interface ReferencedNames {
	projects: Map<string, ReferencedProjectSurface>;
	/** A visible standard module of a referenced project, by name, and the project it is in. */
	modules: Map<string, { project: ReferencedProjectSurface; module: ReferencedProjectModule }>;
	/** Every Enum a referenced project offers, lowercased: which bare `X.Y` to look up in reference order. */
	enumNames: ReadonlySet<string>;
}

function referencedNames(surfaces: readonly ReferencedProjectSurface[]): ReferencedNames {
	const projects = new Map(surfaces.map((surface) => [surface.name.toLowerCase(), surface]));
	const modules = new Map<string, { project: ReferencedProjectSurface; module: ReferencedProjectModule }>();
	for (const project of surfaces) {
		for (const module of referencedProjectNames(project).modules.values()) {
			const key = module.name.toLowerCase();
			if (module.kind === 'standard' && module.visible && !modules.has(key)) {
				modules.set(key, { project, module });
			}
		}
	}
	const enumNames = new Set(surfaces.flatMap((project) => [...referencedProjectNames(project).enums.keys()]));
	return { projects, modules, enumNames };
}

const hasMember = (module: ReferencedProjectModule, name: string): boolean =>
	module.members.some((member) => member.toLowerCase() === name.toLowerCase());

/** A name VBA or the host gives as a value, ahead of any reference: `Strings`, Excel's `ActiveWorkbook`. */
function isVbaOrHostValue(lower: string, model: HostObjectModel | undefined): boolean {
	return resolveRuntimeFunction(lower) !== undefined || resolveRuntimeConstant(lower) !== undefined
		|| resolveRuntimeObject(lower) !== undefined || resolveVbaLibraryQualifier(lower) !== undefined
		|| resolveHostGlobal(lower, model) !== undefined || resolveHostGlobalMember(lower, model) !== undefined
		|| resolveHostConstant(lower, model) !== undefined;
}

export function checkReferencedProjectNames(
	source: string,
	surfaces: readonly ReferencedProjectSurface[] | undefined,
	moduleSymbols: readonly VbaSymbol[],
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectModuleNames: ReadonlySet<string> | undefined,
	order: ReferencedLookupOrder,
	hostModel: HostObjectModel | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	if (!surfaces || surfaces.length === 0) {
		return () => undefined;
	}
	const names = referencedNames(surfaces);
	// What the project itself names wins over a referenced project's name.
	const projectNames = new Set<string>([
		...moduleSymbols.map((symbol) => symbol.name.toLowerCase()),
		...(projectVisibleSymbols ?? []).map((symbol) => symbol.name.toLowerCase()),
		...[...(projectModuleNames ?? [])].map((name) => name.toLowerCase()),
	]);
	const report = (text: string, span: { start: number; end: number }): void => {
		push('memberNotFound', `Method or data member not found: '${text}'.`, span);
	};
	// The qualifiers of names in a type position: `TypeOf o Is ProbeLibrary.Widget`.
	let typeQualifiers: ReadonlySet<number> | undefined;
	const namesAType = (offset: number): boolean => (typeQualifiers ??= new Set(
		collectTypeNameReferences(source).flatMap((ref) => (ref.qualifierSpan ? [ref.qualifierSpan.start] : [])),
	)).has(offset);
	return (procedure) => {
		const own = new Set(procedure.params.map((param) => param.name.toLowerCase()));
		forEachVariableGroup(procedure.body, (group) => {
			for (const decl of group.declarations) {
				own.add(decl.name.toLowerCase());
			}
		});
		const shadowed = (lower: string): boolean => own.has(lower) || projectNames.has(lower);
		return (stmt) => {
			const toks = statementTokens(source, stmt.span);
			for (let i = 0; i + 2 < toks.length; i++) {
				const head = tokenName(toks[i]);
				if (!head || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!' || toks[i + 1].rawText !== '.') {
					continue;
				}
				const second = tokenName(toks[i + 2]);
				if (!second) {
					continue;
				}
				const lower = head.toLowerCase();
				if (shadowed(lower)) {
					continue;
				}
				const project = names.projects.get(lower);
				if (project) {
					// `ReDim a(1) As ProbeLibrary.Widget` names a type too.
					if (toks[i - 1]?.rawText.toLowerCase() === 'as' || namesAType(absoluteSpan(stmt.span, toks[i]).start)) {
						continue;
					}
					const module = referencedProjectNames(project).modules.get(second.toLowerCase());
					if (module?.visible) {
						if (module.kind === 'class' && !module.predeclared) {
							// ProbeLibrary.Widget, a class with no default instance, is no value.
							report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
							continue;
						}
						// ProbeLibrary.Service.Nope(): the module's own members.
						const third = toks[i + 3]?.rawText === '.' ? tokenName(toks[i + 4]) : undefined;
						if (third && module.kind === 'standard' && !hasMember(module, third)) {
							report(`${head}.${second}.${third}`, absoluteSpan(stmt.span, toks[i + 4]));
						}
						continue;
					}
					// ProbeLibrary.WidgetColor.wcRed: an Enum's own members.
					const enumDecl = referencedProjectNames(project).enums.get(second.toLowerCase());
					if (enumDecl) {
						const third = toks[i + 3]?.rawText === '.' ? tokenName(toks[i + 4]) : undefined;
						if (third && !enumDecl.members.some((member) => member.toLowerCase() === third.toLowerCase())) {
							report(`${head}.${second}.${third}`, absoluteSpan(stmt.span, toks[i + 4]));
						}
						continue;
					}
					// ProbeLibrary.Value(): a member of any of its visible standard
					// modules, or ProbeLibrary.wcRed, an exposed class's Enum member.
					const member = referencedProjectNames(project).values.has(second.toLowerCase());
					if (!member) {
						report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
					}
					continue;
				}
				const referenced = names.modules.get(lower);
				// WidgetColor.Nope, where the Enum is what the bare name reaches
				// first: ahead of a later project's module of the name, and behind
				// a value VBA or the host gives, or a library that might have it.
				const reached = names.enumNames.has(lower)
					? referencedProjectReached(lower, surfaces, order, (offered) => offered.bareNames.has(lower) || offered.types.has(lower))
					: undefined;
				const enumDecl = reached && reached !== referenced?.project ? referencedProjectNames(reached).enums.get(lower) : undefined;
				if (enumDecl) {
					if (!isVbaOrHostValue(lower, hostModel) && !enumDecl.members.some((member) => member.toLowerCase() === second.toLowerCase())) {
						report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
					}
					continue;
				}
				// Service.Nope(), Service being the referenced project's.
				if (referenced && !hasMember(referenced.module, second)) {
					report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
				}
			}
		};
	};
}
