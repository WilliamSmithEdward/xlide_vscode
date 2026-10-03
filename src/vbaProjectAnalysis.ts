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
import { logAnalysisFailures } from './analysisFailureLog';

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
    /**
     * The host class the module's designer makes it, where that is not an
     * MSForms.UserForm: an Access form's `Access.Form`. The index hands it to
     * the form's type, so a reference from another module reaches that base.
     */
    designerClass?: string;
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
    | 'hiddenTypeNames'
    | 'projectProcedures'
    | 'projectClassMembers'
    | 'projectTypes'
    | 'projectVisibleSymbols'
    | 'projectIntegerConstants'
    | 'projectStringLiteralWords'
    | 'projectRunnableProcedures'
    | 'projectWrittenNames'
    | 'projectSheetChanges'
    | 'projectOpenedFileNumbers'
    | 'implicitMembers'
    | 'implementedInterfaces'
    | 'conditionalCompilation'
> & {
    /**
     * Why the project-sensitive options are absent, when the index could not
     * answer for the module: the analysis reports it, so a module checked
     * without its project is not taken for a clean one (issue #178).
     */
    projectContextFailure?: unknown;
};

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
        // An Access form or report is the same shape: `Me` is the design, and
        // its controls are members of it.
        case 'accessform':
        case 'accessreport':
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
        designerClass: mod.designerClass,
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

/**
 * A module the live index leaves out, written to the analysis failure log:
 * every other module then sees none of its declarations, and nothing said so
 * (issue #178). A caller that keeps its own record passes onInvalidModule.
 */
function logModuleLeftOut(moduleName: string, error: unknown): void {
    logAnalysisFailures(moduleName, [{
        stage: 'index',
        message: error instanceof Error ? error.message : String(error),
        ...(error instanceof Error && error.stack ? { stack: error.stack } : {}),
    }]);
}

export function buildLiveVbaProjectIndex(
    modules: readonly VbaProjectModuleInput[],
    liveOverride?: VbaProjectLiveOverride,
): ProjectIndex {
    return buildVbaProjectIndex(modules, liveOverride, { ignoreInvalidModules: true, onInvalidModule: logModuleLeftOut });
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
        onInvalidModule: logModuleLeftOut,
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
        // Taken together or not at all: a query failing partway left the
        // answers before it set and the rest absent.
        const answers: VbaProjectAnalysisOptions = {
            knownProcedures: project.visibleProcedureNames(moduleName),
            knownIdentifiers: project.visibleIdentifierNames(moduleName),
            knownNonTypeNames: project.visibleNonTypeNames(moduleName),
            hiddenTypeNames: project.hiddenTypeNames(moduleName),
            projectTypes: project.visibleTypeNames(moduleName),
            projectVisibleSymbols: project.visibleIdentifierSymbols(moduleName),
            projectClassMembers: project.projectMemberSurfaces(moduleName),
            projectIntegerConstants: project.visibleExternalIntegerConstantExpressions(moduleName),
            projectStringLiteralWords: project.stringLiteralWords(),
            projectRunnableProcedures: project.runnableProcedureNames(),
            projectWrittenNames: project.writtenNames(),
            projectSheetChanges: project.sheetChanges(),
            projectOpenedFileNumbers: project.openedFileNumbers(),
            implementedInterfaces: project.implementedInterfaceNames(),
        // The rules must see the same constants the symbol table was built
        // with, or a branch dropped from the symbols would still be analyzed.
            conditionalCompilation: project.conditionalCompilation(),
        };
        // A UserForm's controls are members its own text never declares, so
        // without them every reference in the code-behind reads as undeclared.
        // The index knows them: host-supplied with the module, or parsed from
        // a `.frm` header when the source carries one.
        const controls = project.moduleImplicitMembers?.(moduleName) ?? [];
        if (controls.length > 0) {
            answers.implicitMembers = controls;
        }
        Object.assign(options, answers);
    } catch (err) {
        // Leave every project-sensitive option absent when the index cannot
        // answer the module-specific question. Single-module analysis remains
        // conservative rather than guessing at cross-module visibility, and
        // says so.
        options.projectContextFailure = err;
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
        // Both or neither, as with the analysis options above.
        const procedures = project.visibleProcedureSignatures(moduleName)
            .filter((procedure) => procedure.moduleName.toLowerCase() !== currentLower);
        const symbols = project.visibleIdentifierSymbols(moduleName)
            .filter((symbol) => symbol.moduleName.toLowerCase() !== currentLower);
        externalProjectProcedures = procedures;
        externalProjectSymbols = symbols;
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
