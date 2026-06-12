import { analyzeModule, VbaDiagnostic } from '../../src/analyzer';
import {
	buildVbaProjectIndex,
	projectAnalysisOptionsForModule,
	projectProcedureSignatures,
	type VbaProjectModuleInput,
} from '../../src/vbaProjectAnalysis';

export type ProjectTestModule = VbaProjectModuleInput;

export function projectOptions(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReturnType<typeof projectAnalysisOptionsForModule> {
	const project = buildVbaProjectIndex(modules);
	return projectAnalysisOptionsForModule(
		project,
		currentModule,
		projectProcedureSignatures(project),
	);
}

export function analyzeProjectModule(
	source: string,
	modules: readonly ProjectTestModule[],
	currentModule: string,
	extra: Parameters<typeof analyzeModule>[1] = {},
): VbaDiagnostic[] {
	const current = modules.find(
		(mod) => mod.moduleName.toLowerCase() === currentModule.toLowerCase(),
	);
	const projectModules: readonly ProjectTestModule[] = [
		current
			? { ...current, moduleName: currentModule, source }
			: { moduleName: currentModule, source },
		...modules.filter(
			(mod) => mod.moduleName.toLowerCase() !== currentModule.toLowerCase(),
		),
	];
	return analyzeModule(source, {
		moduleName: currentModule,
		...projectOptions(projectModules, currentModule),
		...extra,
	});
}

export function projectProcedures(
	modules: readonly ProjectTestModule[],
): NonNullable<ReturnType<typeof projectProcedureSignatures>> {
	const project = buildVbaProjectIndex(modules);
	return projectProcedureSignatures(project) ?? new Map();
}

export function projectClassMembers(
	modules: readonly ProjectTestModule[],
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectClassMembers']> {
	const project = buildVbaProjectIndex(modules);
	return project.projectClassMembers();
}

export function projectMemberSurfaces(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectClassMembers']> {
	return projectOptions(modules, currentModule).projectClassMembers ?? [];
}

export function visibleProjectProcedures(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownProcedures ?? new Set();
}

export function visibleProjectIdentifiers(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownIdentifiers ?? new Set();
}

export function visibleProjectTypes(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): NonNullable<ReturnType<typeof projectAnalysisOptionsForModule>['projectTypes']> {
	return projectOptions(modules, currentModule).projectTypes ?? [];
}

export function visibleProjectNonTypeNames(
	modules: readonly ProjectTestModule[],
	currentModule: string,
): ReadonlySet<string> {
	return projectOptions(modules, currentModule).knownNonTypeNames ?? new Set();
}
