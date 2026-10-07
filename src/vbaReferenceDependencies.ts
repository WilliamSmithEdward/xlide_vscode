import {buildVbaProjectIndexAsync,type VbaProjectModuleInput} from './vbaProjectAnalysis';
import {librariesNamedIn} from './analyzer/diagnostics/rules/missingReference';

/** Modules genuinely naming a library, with the same source bindings as diagnostics. */
export async function projectLibraryDependencies(modules: readonly VbaProjectModuleInput[], library: string): Promise<string[]> {
    const project = await buildVbaProjectIndexAsync(modules);
    const names = new Set(modules.map(module => module.moduleName.toLowerCase()));
    const dependencies: string[] = [];
    for (const module of modules) {
        const symbols = project.getModule(module.moduleName);
        if (symbols && librariesNamedIn(module.source, {
            symbols, projectVisibleSymbols: project.visibleIdentifierSymbols(module.moduleName),
        }, names).has(library.toLowerCase())) { dependencies.push(module.moduleName); }
    }
    return dependencies;
}
