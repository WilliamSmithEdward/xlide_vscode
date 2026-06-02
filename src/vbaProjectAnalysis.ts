import {
    ProjectIndex,
    type AnalyzeModuleOptions,
    type EventHandlerDocumentType,
    type ModuleSymbolKind,
} from './analyzer';

export interface VbaProjectModuleInput {
    moduleName: string;
    source: string;
    type?: string;
    moduleKind?: ModuleSymbolKind;
    documentType?: EventHandlerDocumentType;
}

export interface VbaProjectLiveOverride {
    moduleName: string;
    moduleKind: ModuleSymbolKind;
    source: string;
}

export interface VbaProjectIndexBuildOptions {
    ignoreInvalidModules?: boolean;
}

export type VbaProjectAnalysisOptions = Pick<
    AnalyzeModuleOptions,
    | 'knownProcedures'
    | 'knownIdentifiers'
    | 'knownNonTypeNames'
    | 'projectProcedures'
    | 'projectClassMembers'
    | 'projectTypes'
>;

export function moduleKindFromType(type?: string): ModuleSymbolKind {
    switch (type) {
        case 'class': return 'class';
        case 'document': return 'document';
        case 'userform': return 'userform';
        default: return 'standard';
    }
}

export function effectiveModuleKind(input: Pick<VbaProjectModuleInput, 'type' | 'moduleKind'>): ModuleSymbolKind {
    return input.moduleKind ?? moduleKindFromType(input.type);
}

export function buildVbaProjectIndex(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
    options: VbaProjectIndexBuildOptions = {},
): ProjectIndex {
    const index = new ProjectIndex();
    const setModule = (module: Parameters<ProjectIndex['setModule']>[0]): boolean => {
        if (!options.ignoreInvalidModules) {
            index.setModule(module);
            return true;
        }
        try {
            index.setModule(module);
            return true;
        } catch {
            return false;
        }
    };
    let appliedOverride = false;
    for (const mod of modules) {
        const isOverride =
            liveOverride &&
            mod.moduleName.toLowerCase() === liveOverride.moduleName.toLowerCase();
        const applied = setModule({
            moduleName: mod.moduleName,
            moduleKind: isOverride ? liveOverride.moduleKind : effectiveModuleKind(mod),
            source: isOverride ? liveOverride.source : mod.source,
        });
        appliedOverride = appliedOverride || (!!isOverride && applied);
    }
    if (liveOverride && !appliedOverride) {
        setModule(liveOverride);
    }
    return index;
}

export function projectProcedureSignatures(
    project: ProjectIndex,
): ReturnType<ProjectIndex['procedureSignatures']> | undefined {
    try {
        return project.procedureSignatures();
    } catch {
        return undefined;
    }
}

export function projectAnalysisOptionsForModule(
    project: ProjectIndex,
    moduleName: string,
    projectProcedures = projectProcedureSignatures(project),
): VbaProjectAnalysisOptions {
    const options: VbaProjectAnalysisOptions = { projectProcedures };
    try {
        options.knownProcedures = project.visibleProcedureNames(moduleName);
        options.knownIdentifiers = project.visibleIdentifierNames(moduleName);
        options.knownNonTypeNames = project.visibleNonTypeNames(moduleName);
        options.projectTypes = project.visibleTypeNames(moduleName);
        options.projectClassMembers = project.projectMemberSurfaces(moduleName);
    } catch {
        // Leave every project-sensitive option absent when the index cannot
        // answer the module-specific question. Single-module analysis remains
        // conservative rather than guessing at cross-module visibility.
    }
    return options;
}
