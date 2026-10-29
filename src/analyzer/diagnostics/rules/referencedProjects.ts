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

import type { ReferencedProjectModule, ReferencedProjectSurface } from '../../symbols/referencedProjects';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import { absoluteSpan, forEachVariableGroup, statementTokens, tokenName, type ProcedureStatementVisitor } from '../walker';

/** The referenced projects and the standard modules they show, lowercased, for one module's analysis. */
interface ReferencedNames {
	projects: Map<string, ReferencedProjectSurface>;
	/** A visible standard module of a referenced project, by name, and the project it is in. */
	modules: Map<string, { project: ReferencedProjectSurface; module: ReferencedProjectModule }>;
}

function referencedNames(surfaces: readonly ReferencedProjectSurface[]): ReferencedNames {
	const projects = new Map(surfaces.map((surface) => [surface.name.toLowerCase(), surface]));
	const modules = new Map<string, { project: ReferencedProjectSurface; module: ReferencedProjectModule }>();
	for (const project of surfaces) {
		for (const module of project.modules) {
			const key = module.name.toLowerCase();
			if (module.kind === 'standard' && module.visible && !modules.has(key)) {
				modules.set(key, { project, module });
			}
		}
	}
	return { projects, modules };
}

const hasMember = (module: ReferencedProjectModule, name: string): boolean =>
	module.members.some((member) => member.toLowerCase() === name.toLowerCase());

export function checkReferencedProjectNames(
	source: string,
	surfaces: readonly ReferencedProjectSurface[] | undefined,
	moduleSymbols: readonly VbaSymbol[],
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectModuleNames: ReadonlySet<string> | undefined,
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
					const module = project.modules.find((candidate) => candidate.name.toLowerCase() === second.toLowerCase());
					if (module?.visible) {
						// ProbeLibrary.Service.Nope(): the module's own members.
						const third = toks[i + 3]?.rawText === '.' ? tokenName(toks[i + 4]) : undefined;
						if (third && module.kind === 'standard' && !hasMember(module, third)) {
							report(`${head}.${second}.${third}`, absoluteSpan(stmt.span, toks[i + 4]));
						}
						continue;
					}
					// ProbeLibrary.Value(): a member of any of its visible standard modules.
					const member = project.modules.some((candidate) => candidate.kind === 'standard' && candidate.visible && hasMember(candidate, second));
					if (!member) {
						report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
					}
					continue;
				}
				// Service.Nope(), Service being the referenced project's.
				const referenced = names.modules.get(lower);
				if (referenced && !hasMember(referenced.module, second)) {
					report(`${head}.${second}`, absoluteSpan(stmt.span, toks[i + 2]));
				}
			}
		};
	};
}
