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
import { referencedProjectForBareName, referencedProjectNames, type ReferencedProjectModule, type ReferencedProjectSurface } from '../../symbols/referencedProjects';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import { absoluteSpan, forEachVariableGroup, statementTokens, tokenName, type ProcedureStatementVisitor } from '../walker';

/** The referenced projects and the standard modules they show, lowercased, for one module's analysis. */
interface ReferencedNames {
	projects: Map<string, ReferencedProjectSurface>;
	/** A visible standard module of a referenced project, by name, and the project it is in. */
	modules: Map<string, { project: ReferencedProjectSurface; module: ReferencedProjectModule }>;
	/** Every Enum name a referenced project offers: what a bare `X.Y` may need checking for. */
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

/** A name VBA or the host library gives as a value: a function, constant, object, VBA module or host global. */
function vbaOrHostValue(lower: string, model: HostObjectModel | undefined): boolean {
	return resolveRuntimeFunction(lower) !== undefined || resolveRuntimeConstant(lower) !== undefined
		|| resolveRuntimeObject(lower) !== undefined || resolveVbaLibraryQualifier(lower) !== undefined
		|| resolveHostGlobal(lower, model) !== undefined || resolveHostGlobalMember(lower, model) !== undefined
		|| resolveHostConstant(lower, model) !== undefined;
}

const hasMember = (module: ReferencedProjectModule, name: string): boolean =>
	module.members.some((member) => member.toLowerCase() === name.toLowerCase());

export function checkReferencedProjectNames(
	source: string,
	surfaces: readonly ReferencedProjectSurface[] | undefined,
	moduleSymbols: readonly VbaSymbol[],
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectModuleNames: ReadonlySet<string> | undefined,
	order: { libraries: readonly string[] | undefined; hostName: string; hostModel: HostObjectModel | undefined },
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
	// Where a type is named, `Dim w As ProbeLibrary.Widget`: a qualifier's offset.
	let typeQualifiers: ReadonlySet<number> | undefined;
	const inTypePosition = (offset: number): boolean => (typeQualifiers ??= new Set(collectTypeNameReferences(source)
		.flatMap((ref) => ref.qualifierSpan ? [ref.qualifierSpan.start] : []))).has(offset);
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
					if (/^(as|new)$/i.test(toks[i - 1]?.rawText ?? '') || inTypePosition(stmt.span.start + toks[i].start)) {
						continue;
					}
					const module = referencedProjectNames(project).modules.get(second.toLowerCase());
					if (module?.visible) {
						// ProbeLibrary.Service.Nope(): the module's own members.
						const third = toks[i + 3]?.rawText === '.' ? tokenName(toks[i + 4]) : undefined;
						if (module.kind === 'class' && !module.predeclared) {
							// ProbeLibrary.Widget, a class with no default instance, is no value.
							report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
						} else if (third && (module.kind === 'standard' ? !hasMember(module, third) : false)) {
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
				// Service.Nope(), Service being the referenced project's.
				const referenced = names.modules.get(lower);
				if (referenced) {
					if (!hasMember(referenced.module, second)) {
						report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
					}
					continue;
				}
				// WidgetColor.Nope, WidgetColor being the first Enum of the name a
				// bare name reaches (see referencedProjectForBareName). VBA and the
				// host come first with their values too: Excel's ActiveWorkbook.
				if (!names.enumNames.has(lower)) {
					continue;
				}
				const enumDecl = referencedProjectForBareName(lower, surfaces, order.libraries, order.hostName,
					(projectNames) => projectNames.bareNames.has(lower) || projectNames.types.has(lower));
				const decl = enumDecl && referencedProjectNames(enumDecl).enums.get(lower);
				if (decl && !decl.members.some((member) => member.toLowerCase() === second.toLowerCase()) && !vbaOrHostValue(lower, order.hostModel)) {
					report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
				}
			}
		};
	};
}
