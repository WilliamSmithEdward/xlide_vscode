import * as vscode from 'vscode';
import { encodeModuleUri } from './xlideFileSystem';
import { interfacePrefixHits } from './vbaInterfacePrefix';
import {
    collectTypeNameReferences,
    type ProjectIndex,
    type VbaProjectTypeName,
} from './analyzer';
import {
    buildVbaProjectIndex,
    moduleKindFromType,
    type VbaProjectModuleInput,
} from './vbaProjectAnalysis';
import { lineStartOffsets } from './vbaSourceScan';

export { buildVbaProjectIndex, moduleKindFromType };

export type VbaNavigationModule = VbaProjectModuleInput;

/** Converts a 0-based character offset in `source` to a VS Code position. */
export function offsetToPosition(source: string, offset: number): vscode.Position {
    return createOffsetToPositionConverter(source)(offset);
}

/**
 * Returns an offset→position converter that precomputes line starts once and
 * binary-searches per call; build one per module before converting many spans.
 */
export function createOffsetToPositionConverter(
    source: string,
): (offset: number) => vscode.Position {
    const starts = lineStartOffsets(source);
    return (offset) => {
        const limit = Math.min(offset, source.length);
        let lo = 0;
        let hi = starts.length - 1;
        while (lo < hi) {
            const mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= limit) { lo = mid; } else { hi = mid - 1; }
        }
        return new vscode.Position(lo, offset - starts[lo]);
    };
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
    const toPosition = createOffsetToPositionConverter(mod.source);
    return new vscode.Location(
        encodeModuleUri(xlsmPath, mod.moduleName),
        new vscode.Range(toPosition(span.start), toPosition(span.end)),
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
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        const toPosition = createOffsetToPositionConverter(mod.source);
        for (const ref of collectTypeNameReferences(mod.source)) {
            if (ref.name.toLowerCase() !== lower) {
                continue;
            }
            const visibleMatches = typeDefinitionsForReference(project, mod.moduleName, ref);
            if (!visibleMatches.some((definition) => targetKeys.has(typeDefinitionKey(definition)))) {
                continue;
            }
            push(new vscode.Location(
                uri,
                new vscode.Range(toPosition(ref.span.start), toPosition(ref.span.end)),
            ));
        }
    }
    return out;
}

export function typeDefinitionsForReference(
    project: ProjectIndex,
    moduleName: string,
    ref: { name: string; qualifier?: string },
): VbaProjectTypeName[] {
    const definitions = project.resolveTypeDefinitions(moduleName, ref.name);
    if (!ref.qualifier) {
        return definitions;
    }
    const lowerQualifier = ref.qualifier.toLowerCase();
    return definitions.filter(
        (definition) => definition.moduleName.toLowerCase() === lowerQualifier,
    );
}

export function retargetModuleLocation(
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

export function retargetClassModuleLocation(
    location: vscode.Location,
    xlsmPath: string,
    oldName: string,
    newName: string,
): vscode.Location {
    return retargetModuleLocation(location, xlsmPath, oldName, newName);
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
    locations.push(...interfacePrefixLocations(xlsmPath, byModule, project, oldName));
    return newName
        ? locations.map((loc) => retargetClassModuleLocation(loc, xlsmPath, oldName, newName))
        : locations;
}

/**
 * `Private Sub IShape_Draw()` in a class that declares `Implements IShape`.
 *
 * The prefix binds the implementation to the interface, so renaming the
 * interface has to carry it: left behind, the class silently stops
 * implementing anything. Only modules that actually implement the interface
 * are scanned - a name elsewhere that merely starts with the same word is not
 * part of the contract.
 */
function interfacePrefixLocations(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    interfaceName: string,
): vscode.Location[] {
    const wanted = interfaceName.toLowerCase();
    const out: vscode.Location[] = [];
    for (const mod of byModule.values()) {
        const implemented = project.moduleImplementsList?.(mod.moduleName) ?? [];
        if (!implemented.some((name) => name.toLowerCase() === wanted)) {
            continue;
        }
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        for (const hit of interfacePrefixHits(mod.source, interfaceName)) {
            out.push(new vscode.Location(
                uri,
                new vscode.Range(
                    new vscode.Position(hit.line, hit.column),
                    new vscode.Position(hit.line, hit.column + hit.length),
                ),
            ));
        }
    }
    return out;
}
