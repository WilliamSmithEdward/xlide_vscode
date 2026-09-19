// Renaming a UserForm rewrites every use of its name across the project.
//
// A form's name is used two ways. It is a type: `Dim f As UserForm1`,
// `New UserForm1`, `TypeOf x Is UserForm1`. And it is the form's default
// instance, which VBA creates on first use: `UserForm1.Show`,
// `Unload UserForm1`, `With UserForm1`, `Set f = UserForm1`. The instance is a
// bare name anywhere a value can stand, not only a qualifier, so every use of
// the name counts - except a member of some other receiver (`x.UserForm1`,
// `x!UserForm1`), a named argument, a declaration's own name, and a use a
// local or module-level declaration of the same name shadows. Those are
// something else, and rewriting them would corrupt code the rename never
// owned.

import * as vscode from 'vscode';
import { tokenize, type ProjectIndex, type Span } from './analyzer';
import { tokenName } from './analyzer/lexer/tokenHelpers';
import { moduleDocumentUri } from './vbaDocumentLocation';
import {
    offsetToPosition,
    retargetModuleLocation,
    typeReferenceLocations,
    type VbaNavigationModule,
} from './vbaNavigation';
import { shadowScopesFor } from './vbaStandardModuleRename';

/** Keywords whose next name is being declared, not used. */
const DECLARING_KEYWORDS = new Set(['sub', 'function', 'event', 'const', 'type', 'enum']);
/** These declare only after `Property`: `Set UserForm1 = Nothing` is a use. */
const PROPERTY_KINDS = new Set(['get', 'let', 'set']);
const PROPERTY = new Set(['property']);

function isKeyword(token: { kind: string; rawText: string } | undefined, words: ReadonlySet<string>): boolean {
    return token?.kind === 'keyword' && words.has(token.rawText.toLowerCase());
}

/** The spans in `source` where the bare name is the form's default instance. */
function defaultInstanceSpans(source: string, formName: string): Span[] {
    const lower = formName.toLowerCase();
    const shadows = shadowScopesFor(source, lower);
    if (shadows.moduleLevel) {
        return [];
    }
    const tokens = tokenize(source);
    const spans: Span[] = [];
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (tokenName(token)?.toLowerCase() !== lower) {
            continue;
        }
        const prev = tokens[i - 1];
        const next = tokens[i + 1];
        if (prev && (prev.rawText === '.' || prev.rawText === '!')) {
            continue;
        }
        if (next?.rawText === ':=') {
            continue;
        }
        if (isKeyword(prev, DECLARING_KEYWORDS)
            || (isKeyword(prev, PROPERTY_KINDS) && isKeyword(tokens[i - 2], PROPERTY))) {
            continue;
        }
        // A line label: the name alone at the start of a line, then a colon.
        if (next?.kind === 'colon' && (!prev || prev.kind === 'newline')) {
            continue;
        }
        if (shadows.procedures.some((proc) => proc.shadows && token.start >= proc.start && token.start < proc.end)) {
            continue;
        }
        // `[UserForm1]` keeps its brackets; only the name inside changes.
        const bracketed = token.kind === 'bracketedIdentifier' && token.rawText.endsWith(']');
        spans.push({ start: token.start + (bracketed ? 1 : 0), end: token.end - (bracketed ? 1 : 0) });
    }
    return spans;
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

export function projectUserFormReferenceLocations(
    projectPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    newName?: string,
): vscode.Location[] {
    const form = project.getModule(oldName);
    if (form?.moduleKind !== 'userform') {
        return [];
    }
    const lower = oldName.toLowerCase();
    const definitions = project.resolveTypeDefinitions(form.moduleName, oldName).filter(
        (definition) => definition.kind === 'userform' && definition.moduleName.toLowerCase() === lower,
    );
    const locations: vscode.Location[] = [];
    const seen = new Set<string>();
    const add = (location: vscode.Location): void => {
        const key = `${location.uri.toString()}:${location.range.start.line}:${location.range.start.character}`;
        if (!seen.has(key)) {
            seen.add(key);
            locations.push(location);
        }
    };
    for (const location of typeReferenceLocations(projectPath, byModule, project, oldName, definitions, false)) {
        add(location);
    }
    for (const mod of byModule.values()) {
        const uri = moduleDocumentUri(projectPath, mod);
        for (const span of defaultInstanceSpans(mod.source, oldName)) {
            add(new vscode.Location(uri, new vscode.Range(
                offsetToPosition(mod.source, span.start),
                offsetToPosition(mod.source, span.end),
            )));
        }
    }
    const out = newName
        ? locations.map((location) => retargetModuleLocation(location, projectPath, oldName, newName))
        : locations;
    return out.sort(compareLocations);
}

export function projectUserFormReferenceEdit(
    projectPath: string,
    byModule: Map<string, VbaNavigationModule>,
    project: ProjectIndex,
    oldName: string,
    newName: string,
): { edit: vscode.WorkspaceEdit; uris: vscode.Uri[]; count: number } {
    const edit = new vscode.WorkspaceEdit();
    const seenUris = new Map<string, vscode.Uri>();
    let count = 0;
    for (const location of projectUserFormReferenceLocations(projectPath, byModule, project, oldName, newName)) {
        edit.replace(location.uri, location.range, newName);
        seenUris.set(location.uri.toString(), location.uri);
        count++;
    }
    return { edit, uris: [...seenUris.values()], count };
}
