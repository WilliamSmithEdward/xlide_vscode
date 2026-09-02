// VBA navigation providers: document/workspace symbols, definition,
// references, and rename over the shared ProjectIndex.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import {
    XLIDE_SCHEME,
    isVbaDocument,
    moduleIdentityKey,
} from './xlideFileSystem';
import { VbaModuleSymbols } from './vbaSymbolIndex';
import { VBA_IDENTIFIER_RE } from './vbaSourceScan';
import { moduleNamePositionKind } from './vbaModuleNamePosition';
import { recordRename } from './vbaRenameHistory';
import {
    checkRenameName,
    describeRenameCollision,
    findRenameCollision,
} from './vbaRenameValidation';
import {
    ProjectIndex,
    resolveProcedureLabelDefinitionAt,
    resolveTypeReferenceAt,
    type VbaProjectClassMemberDefinition,
    type Span,
    VbaSymbol as AstSymbol,
    type VbaSymbolKind,
} from './analyzer';
import {
    analysisSourceForDocument,
    isStandaloneVbaDocument,
    moduleKindFromDocument,
    moduleNameFromDocument,
} from './vbaDocumentIdentity';
import { moduleDocumentUri, moduleLocationOfDocument, moduleLocationOrThrow } from './vbaDocumentLocation';
import {
    createOffsetToPositionConverter,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeDefinitionsForReference,
    typeReferenceLocations,
} from './vbaNavigation';
import { buildLiveVbaProjectIndexAsync } from './vbaProjectAnalysis';
import {
    collectSymbolReferences,
    projectClassMemberAtDefinition,
    sourceMemberDefinitionsAt,
    type ReferenceSpan,
} from './vbaReferenceResolution';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import {
    documentOutlineSymbolsForSource,
    workspaceSymbols as presentedWorkspaceSymbols,
    type VbaPresentedSymbol,
    type VbaPresentedWorkspaceSymbol,
} from './vbaSymbolPresentation';

function astSymbolKindToVscode(kind: VbaSymbolKind): vscode.SymbolKind {
    switch (kind) {
        case 'module': return vscode.SymbolKind.Module;
        case 'sub': return vscode.SymbolKind.Method;
        case 'function': return vscode.SymbolKind.Function;
        case 'propertyGet':
        case 'propertyLet':
        case 'propertySet':
            return vscode.SymbolKind.Property;
        case 'parameter':
        case 'localVariable':
            return vscode.SymbolKind.Variable;
        case 'moduleVariable':
        case 'typeField':
            return vscode.SymbolKind.Field;
        case 'constant': return vscode.SymbolKind.Constant;
        case 'enum': return vscode.SymbolKind.Enum;
        case 'enumMember': return vscode.SymbolKind.EnumMember;
        case 'type': return vscode.SymbolKind.Struct;
        case 'event': return vscode.SymbolKind.Event;
        case 'declare': return vscode.SymbolKind.Function;
    }
}

type OffsetToPositionConverter = ReturnType<typeof createOffsetToPositionConverter>;

function spanRange(toPosition: OffsetToPositionConverter, span: Span): vscode.Range {
    return new vscode.Range(toPosition(span.start), toPosition(span.end));
}

function documentSymbolToVscode(
    toPosition: OffsetToPositionConverter,
    symbol: VbaPresentedSymbol,
): vscode.DocumentSymbol {
    const out = new vscode.DocumentSymbol(
        symbol.name,
        symbol.detail,
        astSymbolKindToVscode(symbol.kind),
        spanRange(toPosition, symbol.fullSpan),
        spanRange(toPosition, symbol.nameSpan),
    );
    out.children = symbol.children.map((child) => documentSymbolToVscode(toPosition, child));
    return out;
}

function workspaceSymbolToVscode(
    uri: vscode.Uri,
    toPosition: OffsetToPositionConverter,
    symbol: VbaPresentedWorkspaceSymbol,
): vscode.SymbolInformation {
    return new vscode.SymbolInformation(
        symbol.name,
        astSymbolKindToVscode(symbol.kind),
        symbol.containerName,
        new vscode.Location(uri, spanRange(toPosition, symbol.nameSpan)),
    );
}

/** Strip a quoted-string suffix from a module name reference (`'Mod Name'`). */
function unquoteModule(name: string): string {
    if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) {
        return name.slice(1, -1);
    }
    return name;
}

/**
 * Looks at the text immediately preceding the cursor word to detect
 * a qualifier like `ModuleName.` or `'Module Name'.`
 */
function detectQualifier(line: string, wordStart: number): string | undefined {
    // Walk left over `.`
    let i = wordStart;
    if (i <= 0 || line[i - 1] !== '.') { return undefined; }
    i -= 1;
    if (i <= 0) { return undefined; }

    // Quoted module name: 'Some Module'.
    if (line[i - 1] === "'") {
        const closeQuote = i - 1;
        const openQuote = line.lastIndexOf("'", closeQuote - 1);
        if (openQuote >= 0) {
            return line.slice(openQuote + 1, closeQuote);
        }
        return undefined;
    }

    // Plain identifier qualifier.
    const slice = line.slice(0, i);
    const m = /([\p{L}_][\p{L}\p{M}\p{N}_]*)\s*$/u.exec(slice);
    return m?.[1];
}

// ---------------------------------------------------------------------------
// AST project index (Phase 4 wiring): scope-aware definition/references/rename
// ---------------------------------------------------------------------------

/** Translates an AST symbol's nameSpan to a Location in its owning module. */
function astSymbolToLocation(
    projectPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    symbol: AstSymbol,
): vscode.Location | undefined {
    const mod = byModule.get(symbol.moduleName.toLowerCase());
    if (!mod) { return undefined; }
    return new vscode.Location(
        moduleDocumentUri(projectPath, mod),
        new vscode.Range(
            offsetToPosition(mod.source, symbol.nameSpan.start),
            offsetToPosition(mod.source, symbol.nameSpan.end),
        ),
    );
}

function projectMemberDefinitionToLocation(
    projectPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    definition: VbaProjectClassMemberDefinition,
): vscode.Location | undefined {
    const mod = byModule.get(definition.moduleName.toLowerCase());
    if (!mod) { return undefined; }
    return new vscode.Location(
        moduleDocumentUri(projectPath, mod),
        new vscode.Range(
            offsetToPosition(mod.source, definition.nameSpan.start),
            offsetToPosition(mod.source, definition.nameSpan.end),
        ),
    );
}

/** Adapts pure offset-based reference spans into VS Code locations. */
function referenceSpansToLocations(
    projectPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    spans: readonly ReferenceSpan[],
): vscode.Location[] {
    const out: vscode.Location[] = [];
    for (const span of spans) {
        const mod = byModule.get(span.moduleName.toLowerCase());
        if (!mod) { continue; }
        out.push(new vscode.Location(
            moduleDocumentUri(projectPath, mod),
            new vscode.Range(span.line, span.column, span.line, span.column + span.length),
        ));
    }
    return out;
}

function includeTypeDeclaration(
    requested: boolean,
    definitions: readonly { kind: string }[],
): boolean {
    return requested && definitions.some(
        (definition) =>
            definition.kind !== 'class' &&
            definition.kind !== 'document' &&
            definition.kind !== 'userform',
    );
}

// ---------------------------------------------------------------------------
// Provider implementations
// ---------------------------------------------------------------------------

export class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    provideDocumentSymbols(
        document: vscode.TextDocument,
        token: vscode.CancellationToken,
    ): vscode.DocumentSymbol[] {
        if (!isVbaDocument(document)) { return []; }
        if (token.isCancellationRequested) { return []; }
        const source = analysisSourceForDocument(document);
        const moduleName = moduleNameFromDocument(document);
        const toPosition = createOffsetToPositionConverter(source);
        return documentOutlineSymbolsForSource(
            moduleName,
            moduleKindFromDocument(document),
            source,
        ).map((symbol) => documentSymbolToVscode(toPosition, symbol));
    }
}

export class VbaWorkspaceSymbolProvider implements vscode.WorkspaceSymbolProvider {
    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    async provideWorkspaceSymbols(
        query: string,
        token: vscode.CancellationToken,
    ): Promise<vscode.SymbolInformation[]> {
        const out: vscode.SymbolInformation[] = [];
        const projectPaths = new Set<string>();
        for (const document of vscode.workspace.textDocuments) {
            const location = moduleLocationOfDocument(document);
            if (location) {
                projectPaths.add(location.projectPath);
            }
        }

        for (const projectPath of projectPaths) {
            if (token.isCancellationRequested) { return out; }
            try {
                const context = await this._projectIndexService.contextForProject(projectPath, 'live');
                if (token.isCancellationRequested) { return out; }
                const converters = new Map<string, OffsetToPositionConverter>();
                for (const symbol of presentedWorkspaceSymbols(context.project, query)) {
                    const moduleKey = moduleIdentityKey(symbol.moduleName);
                    const mod = context.byModule.get(moduleKey);
                    if (!mod) { continue; }
                    let toPosition = converters.get(moduleKey);
                    if (!toPosition) {
                        toPosition = createOffsetToPositionConverter(mod.source);
                        converters.set(moduleKey, toPosition);
                    }
                    out.push(workspaceSymbolToVscode(
                        moduleDocumentUri(projectPath, mod),
                        toPosition,
                        symbol,
                    ));
                }
            } catch {
                // Workspace symbols are best-effort; skip projects that fail to read.
            }
        }

        for (const document of vscode.workspace.textDocuments) {
            if (token.isCancellationRequested) { return out; }
            if (!isStandaloneVbaDocument(document)) {
                continue;
            }
            const source = analysisSourceForDocument(document);
            const moduleName = moduleNameFromDocument(document);
            const project = await buildLiveVbaProjectIndexAsync([], {
                moduleName,
                moduleKind: moduleKindFromDocument(document),
                source,
            });
            const toPosition = createOffsetToPositionConverter(source);
            for (const symbol of presentedWorkspaceSymbols(project, query)) {
                out.push(workspaceSymbolToVscode(document.uri, toPosition, symbol));
            }
        }

        return out;
    }
}

export class VbaDefinitionProvider implements vscode.DefinitionProvider {
    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    async provideDefinition(
        document: vscode.TextDocument,
        position: vscode.Position,
        token?: vscode.CancellationToken,
    ): Promise<vscode.Location[] | undefined> {
        if (!moduleLocationOfDocument(document)) { return undefined; }
        if (token?.isCancellationRequested) { return undefined; }

        const documentVersion = document.version;
        const source = analysisSourceForDocument(document);
        const offset = document.offsetAt(position);
        const labelDefinition = resolveProcedureLabelDefinitionAt(source, offset);
        if (labelDefinition) {
            return [new vscode.Location(
                document.uri,
                new vscode.Range(
                    document.positionAt(labelDefinition.label.span.start),
                    document.positionAt(labelDefinition.label.span.end),
                ),
            )];
        }

        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const word = document.getText(wordRange);
        const line = document.lineAt(position.line).text;
        const qualifier = detectQualifier(line, wordRange.start.character);

        let projectPath: string;
        let moduleName: string;
        try {
            ({ projectPath, moduleName } = moduleLocationOrThrow(document));
        } catch {
            return undefined;
        }
        const context = await this._projectIndexService.contextForProject(projectPath, 'live');
        if (token?.isCancellationRequested || document.version !== documentVersion) { return undefined; }
        const { modules, project, byModule } = context;
        const current = byModule.get(moduleIdentityKey(moduleName));

        const memberDefinitions = sourceMemberDefinitionsAt(
            source,
            word,
            document.offsetAt(wordRange.end),
            project,
            modules,
            moduleName,
            current?.type,
            current?.documentType,
        );
        if (memberDefinitions.length > 0) {
            const locations = memberDefinitions
                .map((definition) =>
                    projectMemberDefinitionToLocation(projectPath, byModule, definition),
                )
                .filter((loc): loc is vscode.Location => Boolean(loc));
            return locations.length > 0 ? locations : undefined;
        }

        if (!qualifier) {
            const typeReference = resolveTypeReferenceAt(
                source,
                document.offsetAt(position),
                { projectTypes: project.visibleTypeNames(moduleName) },
            );
            if (typeReference) {
                const typeDefinitions = typeDefinitionsForReference(project, moduleName, typeReference);
                const locations = typeDefinitions
                    .map((definition) => projectTypeDefinitionToLocation(projectPath, byModule, definition))
                    .filter((loc): loc is vscode.Location => Boolean(loc));
                if (locations.length > 0) {
                    return locations;
                }
            }
        }

        const defs = qualifier
            ? project.resolveQualifiedDefinition(unquoteModule(qualifier), word)
            : project.resolveDefinition(moduleName, word, offset);

        const locations: vscode.Location[] = [];
        for (const sym of defs) {
            const loc = astSymbolToLocation(projectPath, byModule, sym);
            if (loc) { locations.push(loc); }
        }
        return locations.length > 0 ? locations : undefined;
    }
}

export class VbaReferenceProvider implements vscode.ReferenceProvider {
    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    async provideReferences(
        document: vscode.TextDocument,
        position: vscode.Position,
        context: vscode.ReferenceContext,
        token?: vscode.CancellationToken,
    ): Promise<vscode.Location[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (token?.isCancellationRequested) { return undefined; }
        const documentVersion = document.version;
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        const source = analysisSourceForDocument(document);
        let projectPath: string;
        let moduleName: string;
        try {
            ({ projectPath, moduleName } = moduleLocationOrThrow(document));
        } catch {
            return undefined;
        }
        const navigation = await this._projectIndexService.contextForProject(projectPath, 'live');
        if (token?.isCancellationRequested || document.version !== documentVersion) { return undefined; }
        const { modules, project, byModule } = navigation;
        const current = byModule.get(moduleIdentityKey(moduleName));

        if (position.line === 0 && position.character === 0 && current?.type === 'class') {
            const definitions = project.resolveTypeDefinitions(moduleName, moduleName).filter(
                (definition) =>
                    definition.kind === 'class' &&
                    definition.moduleName.toLowerCase() === moduleName.toLowerCase(),
            );
            if (definitions.length === 0) {
                return undefined;
            }
            return typeReferenceLocations(
                projectPath,
                byModule,
                project,
                moduleName,
                definitions,
                false,
            );
        }

        if (!wordRange) {
            return undefined;
        }

        const word = document.getText(wordRange);

        // A type reference (`As X` / `New X`) takes precedence over bare value
        // resolution, but only when the cursor is not on a member declaration or
        // member-access reference (which the unified resolver owns).
        const onMember = sourceMemberDefinitionsAt(
            source,
            word,
            document.offsetAt(wordRange.end),
            project,
            modules,
            moduleName,
            current?.type,
            current?.documentType,
        ).length > 0
            || (projectClassMemberAtDefinition(
                project,
                moduleName,
                word,
                document.offsetAt(position),
            )?.definitions?.length ?? 0) > 0;
        if (!onMember) {
            const typeReference = resolveTypeReferenceAt(
                source,
                document.offsetAt(position),
                { projectTypes: project.visibleTypeNames(moduleName) },
            );
            if (typeReference) {
                const definitions = typeDefinitionsForReference(project, moduleName, typeReference);
                if (definitions.length > 0) {
                    return typeReferenceLocations(
                        projectPath,
                        byModule,
                        project,
                        typeReference.name,
                        definitions,
                        includeTypeDeclaration(context.includeDeclaration, definitions),
                    );
                }
            }
        }

        return referenceSpansToLocations(
            projectPath,
            byModule,
            collectSymbolReferences(
                byModule,
                project,
                modules,
                source,
                moduleName,
                current,
                word,
                document.offsetAt(wordRange.end),
                document.offsetAt(position),
                context.includeDeclaration,
            ).references,
        );
    }
}

/**
 * Read/write occurrence shading for the identifier under the caret - the
 * first VS Code consumer of the reference KINDS (issue #55): a write shades
 * the way an assignment target should, a read the way a use does.
 */
export class VbaDocumentHighlightProvider implements vscode.DocumentHighlightProvider {
    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    async provideDocumentHighlights(
        document: vscode.TextDocument,
        position: vscode.Position,
        token?: vscode.CancellationToken,
    ): Promise<vscode.DocumentHighlight[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (token?.isCancellationRequested) { return undefined; }
        const documentVersion = document.version;
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const word = document.getText(wordRange);
        const source = analysisSourceForDocument(document);
        let projectPath: string;
        let moduleName: string;
        try {
            ({ projectPath, moduleName } = moduleLocationOrThrow(document));
        } catch {
            return undefined;
        }
        const navigation = await this._projectIndexService.contextForProject(projectPath, 'live');
        if (token?.isCancellationRequested || document.version !== documentVersion) { return undefined; }
        const { modules, project, byModule } = navigation;
        const current = byModule.get(moduleIdentityKey(moduleName));
        const { references } = collectSymbolReferences(
            byModule,
            project,
            modules,
            source,
            moduleName,
            current,
            word,
            document.offsetAt(wordRange.end),
            document.offsetAt(position),
            true,
        );
        const highlights: vscode.DocumentHighlight[] = [];
        for (const span of references) {
            if (span.moduleName.toLowerCase() !== moduleName.toLowerCase()) { continue; }
            highlights.push(new vscode.DocumentHighlight(
                new vscode.Range(span.line, span.column, span.line, span.column + span.length),
                span.kind === 'read'
                    ? vscode.DocumentHighlightKind.Read
                    : vscode.DocumentHighlightKind.Write,
            ));
        }
        return highlights.length > 0 ? highlights : undefined;
    }
}

export class VbaRenameProvider implements vscode.RenameProvider {
    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    async prepareRename(
        document: vscode.TextDocument,
        position: vscode.Position,
        token?: vscode.CancellationToken,
    ): Promise<vscode.Range | { range: vscode.Range; placeholder: string }> {
        if (document.uri.scheme !== XLIDE_SCHEME) {
            throw new Error('Rename is only supported in XLIDE VBA modules.');
        }
        if (token?.isCancellationRequested) {
            throw new vscode.CancellationError();
        }
        const documentVersion = document.version;
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { throw new Error('No symbol at cursor.'); }
        const word = document.getText(wordRange);

        const source = analysisSourceForDocument(document);
        let projectPath: string;
        let moduleName: string;
        try {
            ({ projectPath, moduleName } = moduleLocationOrThrow(document));
        } catch {
            throw new Error('XLIDE cannot rename here: this is not a project VBA module.');
        }
        const navigation = await this._projectIndexService.contextForProject(projectPath, 'strict');
        if (token?.isCancellationRequested || document.version !== documentVersion) {
            throw new vscode.CancellationError();
        }
        const { modules, project, byModule } = navigation;
        const current = byModule.get(moduleIdentityKey(moduleName));

        const memberDefinitions = sourceMemberDefinitionsAt(
            source,
            word,
            document.offsetAt(wordRange.end),
            project,
            modules,
            moduleName,
            current?.type,
            current?.documentType,
        );
        const memberAtDefinition = memberDefinitions.length > 0
            ? undefined
            : projectClassMemberAtDefinition(
                project,
                moduleName,
                word,
                document.offsetAt(position),
            );
        if (memberDefinitions.length > 0 || (memberAtDefinition?.definitions?.length ?? 0) > 0) {
            return { range: wordRange, placeholder: word };
        }

        const scope = project.referenceScope(
            moduleName,
            word,
            document.offsetAt(position),
        );
        if (scope.definitions.length === 0) {
            throw new Error(`'${word}' is not a renameable VBA symbol in this project.`);
        }
        return { range: wordRange, placeholder: word };
    }

    async provideRenameEdits(
        document: vscode.TextDocument,
        position: vscode.Position,
        newName: string,
        token?: vscode.CancellationToken,
    ): Promise<vscode.WorkspaceEdit | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (token?.isCancellationRequested) {
            return undefined;
        }
        const documentVersion = document.version;
        const nameProblem = checkRenameName(newName);
        if (nameProblem) {
            throw new Error(nameProblem.message);
        }
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const oldName = document.getText(wordRange);
        // Only short-circuit a byte-identical rename. A case-only change (myVar ->
        // MyVar) is a legitimate operation for a canonical-case tool, not a no-op.
        if (oldName === newName) { return undefined; }

        const source = analysisSourceForDocument(document);
        let projectPath: string;
        let moduleName: string;
        try {
            ({ projectPath, moduleName } = moduleLocationOrThrow(document));
        } catch {
            throw new Error('XLIDE cannot rename here: this is not a project VBA module.');
        }
        const navigation = await this._projectIndexService.contextForProject(projectPath, 'strict');
        if (token?.isCancellationRequested || document.version !== documentVersion) {
            return undefined;
        }
        const { modules, project, byModule } = navigation;
        const current = byModule.get(moduleIdentityKey(moduleName));

        const result = collectSymbolReferences(
            byModule,
            project,
            modules,
            source,
            moduleName,
            current,
            oldName,
            document.offsetAt(wordRange.end),
            document.offsetAt(position),
            true,
        );
        if (!result.hasSymbol) {
            // Rule 3: nothing else resolved here, so if the word stands where
            // only a module can, it IS the module's name - which lives on the
            // component rather than in any module's text, so this provider
            // cannot rename it. Say which operation does, instead of implying
            // the name means nothing.
            const asModule = moduleNamePositionKind(
                source,
                document.offsetAt(wordRange.start),
                document.offsetAt(wordRange.end),
            );
            if (asModule) {
                throw new Error(
                    `'${oldName}' is a module name here, and a module's name belongs to the `
                    + 'component rather than to the text of any module. Use Rename Module in the '
                    + 'XLIDE Explorer: it renames the component and updates every reference, '
                    + 'including Implements and Interface_Member prefixes.',
                );
            }
            throw new Error(`'${oldName}' is not a renameable VBA symbol in this project.`);
        }

        // Nothing is written until the new name is known to be free where the
        // old one is declared. Renaming Alpha to Beta beside an existing Beta
        // produced two `Public Sub Beta()` in one module: a project that no
        // longer compiles, reported as a rename that worked.
        for (const editedModule of renamedModuleNames(result.references)) {
            const collision = findRenameCollision(
                project.documentSymbols(editedModule),
                editedModule,
                oldName,
                newName,
            );
            if (collision) {
                throw new Error(describeRenameCollision(collision, newName));
            }
        }

        if (result.ambiguous.length > 0) {
            const where = result.ambiguous
                .slice(0, 5)
                .map((span) => `${span.moduleName}:${span.line + 1}`)
                .join(', ');
            const more = result.ambiguous.length > 5 ? `, and ${result.ambiguous.length - 5} more` : '';
            void vscode.window.showWarningMessage(
                `Renamed '${oldName}' to '${newName}', but left ${result.ambiguous.length} `
                + `unqualified reference${result.ambiguous.length === 1 ? '' : 's'} alone because `
                + `another module exports the same name and nothing proves which was meant: `
                + `${where}${more}. Qualify those calls, or check them by hand.`,
            );
        }

        // Rule 10: an editor's undo stack is per document, so undoing in the
        // file you are looking at would reverse that file's share and leave
        // the rest renamed. Keep what each module said immediately before the
        // write, and offer to put all of them back together.
        recordRename({
            projectPath: projectPath,
            oldName,
            newName,
            modules: renamedModuleNames(result.references)
                .map((moduleName) => ({
                    moduleName,
                    before: byModule.get(moduleIdentityKey(moduleName))?.source ?? '',
                }))
                .filter((image) => image.before !== ''),
        });

        const edit = new vscode.WorkspaceEdit();
        for (const loc of referenceSpansToLocations(projectPath, byModule, result.references)) {
            edit.replace(loc.uri, loc.range, newName);
        }
        const touched = renamedModuleNames(result.references).length;
        if (touched > 1) {
            // Only worth offering when the rename spanned modules: that is
            // exactly when per-document undo would leave the project half
            // renamed.
            void vscode.window.showInformationMessage(
                `Renamed '${oldName}' to '${newName}' across ${touched} modules.`,
                'Undo Rename',
            ).then((choice) => {
                if (choice === 'Undo Rename') {
                    void vscode.commands.executeCommand('xlide.undoRename');
                }
            });
        }
        return edit;
    }
}

function renamedModuleNames(references: readonly { moduleName: string }[]): string[] {
    const seen = new Map<string, string>();
    for (const reference of references) {
        const key = reference.moduleName.toLowerCase();
        if (!seen.has(key)) {
            seen.set(key, reference.moduleName);
        }
    }
    return [...seen.values()];
}
