import * as vscode from 'vscode';
import { encodeModuleUri } from './xlideFileSystem';
import {
    ProjectIndex,
    resolveTypeSemanticTokens,
    type EventHandlerDocumentType,
    type ModuleSymbolKind,
    type VbaProjectTypeName,
} from './analyzer';

export interface VbaNavigationModule {
    moduleName: string;
    source: string;
    type?: string;
    documentType?: EventHandlerDocumentType;
}

export function moduleKindFromType(type?: string): ModuleSymbolKind {
    switch (type) {
        case 'class': return 'class';
        case 'document': return 'document';
        case 'userform': return 'userform';
        default: return 'standard';
    }
}

export function buildVbaProjectIndex(
    modules: readonly VbaNavigationModule[],
    liveOverride?: { moduleName: string; moduleKind: ModuleSymbolKind; source: string },
): ProjectIndex {
    const index = new ProjectIndex();
    let appliedOverride = false;
    for (const mod of modules) {
        const isOverride =
            liveOverride &&
            mod.moduleName.toLowerCase() === liveOverride.moduleName.toLowerCase();
        index.setModule({
            moduleName: mod.moduleName,
            moduleKind: isOverride ? liveOverride.moduleKind : moduleKindFromType(mod.type),
            source: isOverride ? liveOverride.source : mod.source,
        });
        appliedOverride = appliedOverride || !!isOverride;
    }
    if (liveOverride && !appliedOverride) {
        index.setModule(liveOverride);
    }
    return index;
}

/** Converts a 0-based character offset in `source` to a VS Code position. */
export function offsetToPosition(source: string, offset: number): vscode.Position {
    let line = 0;
    let lineStart = 0;
    const limit = Math.min(offset, source.length);
    for (let i = 0; i < limit; i++) {
        if (source[i] === '\n') {
            line++;
            lineStart = i + 1;
        }
    }
    return new vscode.Position(line, offset - lineStart);
}

function typeDefinitionKey(definition: VbaProjectTypeName): string {
    return `${definition.moduleName.toLowerCase()}:${definition.kind}:${definition.nameSpan?.start ?? 0}`;
}

export function projectTypeDefinitionToLocation(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    definition: VbaProjectTypeName,
): vscode.Location | undefined {
    const mod = byModule.get(definition.moduleName.toLowerCase());
    if (!mod) { return undefined; }
    const span = definition.nameSpan ?? { start: 0, end: 0 };
    return new vscode.Location(
        encodeModuleUri(xlsmPath, mod.moduleName),
        new vscode.Range(
            offsetToPosition(mod.source, span.start),
            offsetToPosition(mod.source, span.end),
        ),
    );
}

export function projectClassModuleDefinition(
    project: ProjectIndex,
    moduleName: string,
    className: string,
): VbaProjectTypeName | undefined {
    const lower = className.toLowerCase();
    const definitions = project.resolveTypeDefinitions(moduleName, className).filter(
        (definition) =>
            definition.kind === 'class' &&
            definition.moduleName.toLowerCase() === lower,
    );
    return definitions.length === 1 ? definitions[0] : undefined;
}

export function typeReferenceLocations(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    typeName: string,
    definitions: readonly VbaProjectTypeName[],
    includeDeclaration: boolean,
): vscode.Location[] {
    const lower = typeName.toLowerCase();
    const targetKeys = new Set(definitions.map(typeDefinitionKey));
    const seen = new Set<string>();
    const out: vscode.Location[] = [];
    const push = (location: vscode.Location): void => {
        const key = `${location.uri.toString()}:${location.range.start.line}:${location.range.start.character}`;
        if (!seen.has(key)) {
            seen.add(key);
            out.push(location);
        }
    };

    if (includeDeclaration) {
        for (const definition of definitions) {
            const loc = projectTypeDefinitionToLocation(xlsmPath, byModule, definition);
            if (loc) { push(loc); }
        }
    }

    for (const mod of byModule.values()) {
        const visibleMatches = project.resolveTypeDefinitions(mod.moduleName, typeName);
        if (!visibleMatches.some((definition) => targetKeys.has(typeDefinitionKey(definition)))) {
            continue;
        }
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        for (const token of resolveTypeSemanticTokens(mod.source, {
            projectTypes: project.visibleTypeNames(mod.moduleName),
        })) {
            if (token.name.toLowerCase() !== lower) {
                continue;
            }
            push(new vscode.Location(
                uri,
                new vscode.Range(
                    offsetToPosition(mod.source, token.span.start),
                    offsetToPosition(mod.source, token.span.end),
                ),
            ));
        }
    }
    return out;
}

export function retargetClassModuleLocation(
    location: vscode.Location,
    xlsmPath: string,
    oldName: string,
    newName: string,
): vscode.Location {
    const oldUri = encodeModuleUri(xlsmPath, oldName).toString();
    if (location.uri.toString() !== oldUri) {
        return location;
    }
    return new vscode.Location(
        encodeModuleUri(xlsmPath, newName),
        location.range,
    );
}

export function projectClassReferenceLocations(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    definition: VbaProjectTypeName,
    newName?: string,
): vscode.Location[] {
    const locations = typeReferenceLocations(
        xlsmPath,
        byModule,
        project,
        oldName,
        [definition],
        false,
    );
    return newName
        ? locations.map((loc) => retargetClassModuleLocation(loc, xlsmPath, oldName, newName))
        : locations;
}
