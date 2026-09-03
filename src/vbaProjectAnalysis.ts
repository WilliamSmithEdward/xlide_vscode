import {
    ProjectIndex,
    type AnalyzeModuleOptions,
    type ConditionalCompilationEnvironment,
    type EventHandlerDocumentType,
    type ModuleSymbolKind,
    type VbaProcedureSignature,
    type VbaSymbol,
} from './analyzer';
import { yieldToExtensionHost } from './util/async';

export interface VbaProjectModuleInput {
    moduleName: string;
    source: string;
    type?: string;
    moduleKind?: ModuleSymbolKind;
    documentType?: EventHandlerDocumentType;
    /**
     * A form's designer-declared controls, from a host that can read the
     * designer. The index folds them into the form's member surface so a
     * qualified reference from another module resolves them.
     */
    implicitMembers?: readonly { name: string; type: string }[];
    /**
     * True when the module carries `Attribute VB_PredeclaredId = True`, giving
     * it a default instance so its own name is usable as a value. Absent means
     * the attribute header was not read, never "no".
     */
    predeclaredId?: boolean;
}

export interface VbaProjectLiveOverride {
    moduleName: string;
    moduleKind: ModuleSymbolKind;
    source: string;
}

export interface VbaProjectIndexBuildOptions {
    ignoreInvalidModules?: boolean;
    /** Optional cooperative cancellation hook used by async callers. */
    cancelIfRequested?: () => void;
    /** Test/host override for how often the async builder yields. */
    yieldEveryModules?: number;
    /** Called for each module skipped by ignoreInvalidModules. */
    onInvalidModule?: (moduleName: string, error: unknown) => void;
    /**
     * The project's own conditional compilation arguments, from the VBE project
     * property. Supplying them lets `#If MY_FLAG` be decided instead of leaving
     * every arm live.
     */
    conditionalCompilation?: ConditionalCompilationEnvironment;
}

const PROJECT_INDEX_YIELD_EVERY_MODULES = 8;

export type VbaProjectAnalysisOptions = Pick<
    AnalyzeModuleOptions,
    | 'knownProcedures'
    | 'knownIdentifiers'
    | 'knownNonTypeNames'
    | 'projectProcedures'
    | 'projectClassMembers'
    | 'projectTypes'
    | 'projectVisibleSymbols'
    | 'projectIntegerConstants'
    | 'implicitMembers'
    | 'implementedInterfaces'
    | 'conditionalCompilation'
>;

export interface VbaProjectEditorSymbolContext {
    analysisOptions: VbaProjectAnalysisOptions;
    externalProjectProcedures: VbaProcedureSignature[];
    externalProjectSymbols: VbaSymbol[];
}

export function moduleKindFromType(type?: string): ModuleSymbolKind {
    switch (type) {
        case 'class': return 'class';
        case 'document': return 'document';
        case 'userform': return 'userform';
        // A VB6 UserControl, PropertyPage or Designer is an object module
        // with a designer, like a form: `Me` is valid and controls live on it.
        case 'usercontrol':
        case 'propertypage':
        case 'designer':
            return 'userform';
        default: return 'standard';
    }
}

export function effectiveModuleKind(input: Pick<VbaProjectModuleInput, 'type' | 'moduleKind'>): ModuleSymbolKind {
    return input.moduleKind ?? moduleKindFromType(input.type);
}

type ProjectIndexModuleSetter = (module: Parameters<ProjectIndex['setModule']>[0]) => boolean;

function projectIndexModuleSetter(
    index: ProjectIndex,
    options: VbaProjectIndexBuildOptions,
): ProjectIndexModuleSetter {
    return (module) => {
        if (!options.ignoreInvalidModules) {
            index.setModule(module);
            return true;
        }
        try {
            index.setModule(module);
            return true;
        } catch (err) {
            options.onInvalidModule?.(module.moduleName, err);
            return false;
        }
    };
}

/** Sets one module on the index; returns whether it consumed the live override. */
function applyProjectModule(
    setModule: ProjectIndexModuleSetter,
    mod: VbaProjectModuleInput,
    liveOverride?: VbaProjectLiveOverride,
): boolean {
    const isOverride =
        liveOverride &&
        mod.moduleName.toLowerCase() === liveOverride.moduleName.toLowerCase();
    const applied = setModule({
        moduleName: mod.moduleName,
        moduleKind: isOverride ? liveOverride.moduleKind : effectiveModuleKind(mod),
        source: isOverride ? liveOverride.source : mod.source,
        implicitMembers: mod.implicitMembers,
        predeclaredId: mod.predeclaredId,
    });
    return !!isOverride && applied;
}

export function buildVbaProjectIndex(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
    options: VbaProjectIndexBuildOptions = {},
): ProjectIndex {
    const index = new ProjectIndex({ conditionalCompilation: options.conditionalCompilation });
    const setModule = projectIndexModuleSetter(index, options);
    let appliedOverride = false;
    for (const mod of modules) {
        appliedOverride = applyProjectModule(setModule, mod, liveOverride) || appliedOverride;
    }
    if (liveOverride && !appliedOverride) {
        setModule(liveOverride);
    }
    return index;
}

export function buildLiveVbaProjectIndex(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
): ProjectIndex {
    return buildVbaProjectIndex(modules, liveOverride, { ignoreInvalidModules: true });
}

export async function buildVbaProjectIndexAsync(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
    options: VbaProjectIndexBuildOptions = {},
): Promise<ProjectIndex> {
    const index = new ProjectIndex({ conditionalCompilation: options.conditionalCompilation });
    const setModule = projectIndexModuleSetter(index, options);
    let appliedOverride = false;
    const yieldEvery = Math.max(1, options.yieldEveryModules ?? PROJECT_INDEX_YIELD_EVERY_MODULES);
    options.cancelIfRequested?.();
    for (const [i, mod] of modules.entries()) {
        appliedOverride = applyProjectModule(setModule, mod, liveOverride) || appliedOverride;
        if ((i + 1) % yieldEvery === 0) {
            await yieldToExtensionHost();
            options.cancelIfRequested?.();
        }
    }
    if (liveOverride && !appliedOverride) {
        options.cancelIfRequested?.();
        setModule(liveOverride);
    }
    return index;
}

export function buildLiveVbaProjectIndexAsync(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
    options: Omit<VbaProjectIndexBuildOptions, 'ignoreInvalidModules'> = {},
): Promise<ProjectIndex> {
    return buildVbaProjectIndexAsync(modules, liveOverride, {
        ...options,
        ignoreInvalidModules: true,
    });
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
        options.projectVisibleSymbols = project.visibleIdentifierSymbols(moduleName);
        options.projectClassMembers = project.projectMemberSurfaces(moduleName);
        options.projectIntegerConstants = project.visibleExternalIntegerConstantExpressions(moduleName);
        options.implementedInterfaces = project.implementedInterfaceNames();
        // The rules must see the same constants the symbol table was built
        // with, or a branch dropped from the symbols would still be analyzed.
        options.conditionalCompilation = project.conditionalCompilation();
        // A UserForm's controls are members its own text never declares, so
        // without them every reference in the code-behind reads as undeclared.
        // The index knows them: host-supplied with the module, or parsed from
        // a `.frm` header when the source carries one.
        const controls = project.moduleImplicitMembers?.(moduleName) ?? [];
        if (controls.length > 0) {
            options.implicitMembers = controls;
        }
    } catch {
        // Leave every project-sensitive option absent when the index cannot
        // answer the module-specific question. Single-module analysis remains
        // conservative rather than guessing at cross-module visibility.
    }
    return options;
}

export function projectEditorSymbolContextForModule(
    project: ProjectIndex,
    moduleName: string,
): VbaProjectEditorSymbolContext {
    const analysisOptions = projectAnalysisOptionsForModule(project, moduleName);
    const currentLower = moduleName.toLowerCase();
    let externalProjectProcedures: VbaProcedureSignature[] = [];
    let externalProjectSymbols: VbaSymbol[] = [];
    try {
        externalProjectProcedures = project.visibleProcedureSignatures(moduleName)
            .filter((procedure) => procedure.moduleName.toLowerCase() !== currentLower);
        externalProjectSymbols = project.visibleIdentifierSymbols(moduleName)
            .filter((symbol) => symbol.moduleName.toLowerCase() !== currentLower);
    } catch {
        // Keep project editor surfaces conservative if the index cannot answer
        // visibility for this module.
    }
    return {
        analysisOptions,
        externalProjectProcedures,
        externalProjectSymbols,
    };
}
