import * as vscode from 'vscode';
import {
    collectTypeNameReferences,
    parseModule,
    tokenize,
    type ProjectIndex,
    type Span,
} from './analyzer';
import { WorkbookEngine } from './workbookEngine';
import { encodeModuleUri, notifySignatureDropped } from './xlideFileSystem';
import { moduleDocumentUri } from './vbaDocumentLocation';
import {
    offsetToPosition,
    retargetModuleLocation,
    typeDefinitionsForReference,
    type VbaNavigationModule,
} from './vbaNavigation';

type StandardModuleReferenceEdit = {
    edit: vscode.WorkspaceEdit;
    uris: vscode.Uri[];
    count: number;
};

export async function renameProjectStandardModule(
    bridge: WorkbookEngine,
    xlsmPath: string,
    oldName: string,
    newName: string,
): Promise<void> {
    const result = await bridge.call<{ ok: boolean; signatureDropped: boolean }>(
        'renameModule',
        { path: xlsmPath, module: oldName, newName },
    );
    notifySignatureDropped(xlsmPath, result.signatureDropped);
}

function tokenName(token: ReturnType<typeof tokenize>[number] | undefined): string | undefined {
    if (!token) {
        return undefined;
    }
    if (token.kind === 'identifier' || token.kind === 'keyword') {
        return token.rawText;
    }
    if (token.kind === 'bracketedIdentifier') {
        return token.rawText.slice(1, -1);
    }
    return undefined;
}

function spanLocation(
    xlsmPath: string,
    mod: VbaNavigationModule,
    span: Span,
): vscode.Location {
    return new vscode.Location(
        moduleDocumentUri(xlsmPath, mod),
        new vscode.Range(
            offsetToPosition(mod.source, span.start),
            offsetToPosition(mod.source, span.end),
        ),
    );
}

function isVisibleStandardModuleMemberReference(
    project: ProjectIndex,
    currentModuleName: string,
    moduleName: string,
    memberName: string,
): boolean {
    const lowerModule = moduleName.toLowerCase();
    const surface = project.projectStandardModuleMembers(currentModuleName).find(
        (candidate) => candidate.moduleName.toLowerCase() === lowerModule,
    );
    return (surface?.members ?? []).some(
        (member) => member.name.toLowerCase() === memberName.toLowerCase(),
    );
}

function addLocation(
    locations: vscode.Location[],
    seen: Set<string>,
    location: vscode.Location,
): void {
    const key = `${location.uri.toString()}:${location.range.start.line}:${location.range.start.character}`;
    if (seen.has(key)) {
        return;
    }
    seen.add(key);
    locations.push(location);
}

// Scopes in which a declaration shadows the module name: within them,
// `OldName.Member` refers to the shadowing variable or parameter, not the
// module, so the rename must leave those occurrences alone.
interface ModuleNameShadowScopes {
    /** A module-level declaration shadows the name for the whole module. */
    moduleLevel: boolean;
    procedures: { start: number; end: number; shadows: boolean }[];
}

function shadowScopesFor(source: string, lowerOld: string): ModuleNameShadowScopes {
    const parsed = parseModule(source);
    type BodyLike = { kind?: string; declarations?: { name: string }[]; body?: unknown };
    const bodyDeclares = (body: readonly unknown[]): boolean => {
        for (const node of body as readonly BodyLike[]) {
            if (node.kind === 'VariableGroup') {
                if ((node.declarations ?? []).some((decl) => decl.name.toLowerCase() === lowerOld)) {
                    return true;
                }
            } else if (Array.isArray(node.body) && bodyDeclares(node.body)) {
                return true;
            }
        }
        return false;
    };
    let moduleLevel = false;
    const procedures: ModuleNameShadowScopes['procedures'] = [];
    for (const member of parsed.members) {
        if (member.kind === 'VariableGroup') {
            if (member.declarations.some((decl) => decl.name.toLowerCase() === lowerOld)) {
                moduleLevel = true;
            }
        } else if (member.kind === 'Procedure') {
            const shadows = member.params.some((param) => param.name.toLowerCase() === lowerOld)
                || bodyDeclares(member.body);
            procedures.push({ start: member.span.start, end: member.span.end, shadows });
        }
    }
    return { moduleLevel, procedures };
}

function addQualifiedMemberQualifierLocations(
    xlsmPath: string,
    mod: VbaNavigationModule,
    project: ProjectIndex,
    oldName: string,
    locations: vscode.Location[],
    seen: Set<string>,
): void {
    const lowerOld = oldName.toLowerCase();
    const shadowScopes = shadowScopesFor(mod.source, lowerOld);
    if (shadowScopes.moduleLevel) {
        return;
    }
    const tokens = tokenize(mod.source);
    for (let i = 0; i + 2 < tokens.length; i++) {
        const qualifierName = tokenName(tokens[i]);
        if (!qualifierName || qualifierName.toLowerCase() !== lowerOld) {
            continue;
        }
        // `rs.Fields` / `rs!Fields` / a With block's `.Fields`: the name is a
        // member of some other receiver here, not the module qualifier.
        const prev = tokens[i - 1];
        if (prev && (prev.rawText === '.' || prev.rawText === '!')) {
            continue;
        }
        if (tokens[i + 1].rawText !== '.') {
            continue;
        }
        const memberName = tokenName(tokens[i + 2]);
        if (!memberName) {
            continue;
        }
        const start = tokens[i].start;
        if (shadowScopes.procedures.some(
            (proc) => proc.shadows && start >= proc.start && start < proc.end,
        )) {
            continue;
        }
        if (!isVisibleStandardModuleMemberReference(project, mod.moduleName, oldName, memberName)) {
            continue;
        }
        addLocation(
            locations,
            seen,
            spanLocation(xlsmPath, mod, { start: tokens[i].start, end: tokens[i].end }),
        );
    }
}

function addQualifiedTypeQualifierLocations(
    xlsmPath: string,
    mod: VbaNavigationModule,
    project: ProjectIndex,
    oldName: string,
    locations: vscode.Location[],
    seen: Set<string>,
): void {
    const lowerOld = oldName.toLowerCase();
    for (const ref of collectTypeNameReferences(mod.source)) {
        if (!ref.qualifierSpan || ref.qualifier?.toLowerCase() !== lowerOld) {
            continue;
        }
        const definitions = typeDefinitionsForReference(project, mod.moduleName, ref);
        if (!definitions.some((definition) => definition.moduleName.toLowerCase() === lowerOld)) {
            continue;
        }
        addLocation(locations, seen, spanLocation(xlsmPath, mod, ref.qualifierSpan));
    }
}

function compareLocations(a: vscode.Location, b: vscode.Location): number {
    const uriCmp = a.uri.toString().localeCompare(b.uri.toString());
    if (uriCmp !== 0) {
        return uriCmp;
    }
    if (a.range.start.line !== b.range.start.line) {
        return a.range.start.line - b.range.start.line;
    }
    return a.range.start.character - b.range.start.character;
}

export function projectStandardModuleReferenceLocations(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    newName?: string,
): vscode.Location[] {
    const oldModule = project.getModule(oldName);
    if (oldModule?.moduleKind !== 'standard') {
        return [];
    }

    const locations: vscode.Location[] = [];
    const seen = new Set<string>();
    for (const mod of byModule.values()) {
        addQualifiedMemberQualifierLocations(xlsmPath, mod, project, oldName, locations, seen);
        addQualifiedTypeQualifierLocations(xlsmPath, mod, project, oldName, locations, seen);
    }

    const out = newName
        ? locations.map((location) => retargetModuleLocation(location, xlsmPath, oldName, newName))
        : locations;
    return out.sort(compareLocations);
}

export function projectStandardModuleReferenceEdit(
    xlsmPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    newName: string,
): StandardModuleReferenceEdit {
    const edit = new vscode.WorkspaceEdit();
    const seenUris = new Map<string, vscode.Uri>();
    let count = 0;
    for (const location of projectStandardModuleReferenceLocations(
        xlsmPath,
        byModule,
        project,
        oldName,
        newName,
    )) {
        edit.replace(location.uri, location.range, newName);
        seenUris.set(location.uri.toString(), location.uri);
        count++;
    }
    return { edit, uris: [...seenUris.values()], count };
}
