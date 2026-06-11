// VBA navigation providers: document/workspace symbols, definition,
// references, and rename over the shared workbook ProjectIndex.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import {
    XLIDE_SCHEME,
    decodeModuleUri,
    encodeModuleUri,
    isVbaDocument,
    moduleIdentityKey,
} from './xlideFileSystem';
import { VbaModuleSymbols } from './vbaSymbolIndex';
import {
    findIdentifierOccurrences,
    lineStartOffsets,
    VBA_IDENTIFIER_NAME_RE,
    VBA_IDENTIFIER_RE,
} from './vbaSourceScan';
import {
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    ProjectIndex,
    ReferenceScope,
    resolveMemberDefinitionsAt,
    resolveProcedureLabelDefinitionAt,
    tokenizeCached,
    type MemberCompletionContext,
    resolveTypeReferenceAt,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    type Span,
    VbaSymbol as AstSymbol,
    type VbaSymbolKind,
} from './analyzer';
import {
    isStandaloneVbaDocument,
    moduleKindFromDocument,
    moduleNameFromDocument,
} from './vbaDocumentIdentity';
import {
    createOffsetToPositionConverter,
    moduleKindFromType,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeDefinitionsForReference,
    typeReferenceLocations,
} from './vbaNavigation';
import { buildLiveVbaProjectIndexAsync } from './vbaProjectAnalysis';
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
    const m = /([A-Za-z_][A-Za-z0-9_]*)\s*$/.exec(slice);
    return m?.[1];
}

// ---------------------------------------------------------------------------
// AST project index (Phase 4 wiring): scope-aware definition/references/rename
// ---------------------------------------------------------------------------

function codeNamesForModules(modules: VbaModuleSymbols[]): Record<string, string> {
    const out: Record<string, string> = {};
    for (const mod of modules) {
        const hostType = documentHostTypeForModule(mod.moduleName, mod.type, mod.documentType);
        if (!hostType) { continue; }
        out[mod.moduleName.toLowerCase()] = hostType;
    }
    return out;
}

function meProjectTypeForModule(moduleName: string, type?: string): string | undefined {
    const kind = moduleKindFromType(type);
    return kind === 'class' || kind === 'document' || kind === 'userform'
        ? moduleName
        : undefined;
}

function meHostTypeForModule(
    moduleName: string,
    type?: string,
    documentType?: EventHandlerDocumentType,
): string | undefined {
    return documentHostTypeForModule(moduleName, type, documentType);
}

function documentHostTypeForModule(
    moduleName: string,
    type?: string,
    documentType?: EventHandlerDocumentType,
): string | undefined {
    switch (eventHandlerDocumentTypeForContext({
        moduleName,
        moduleKind: moduleKindFromType(type),
        documentType,
    })) {
        case 'workbook':
            return 'Excel.Workbook';
        case 'chart':
            return 'Excel.Chart';
        case 'worksheet':
            return 'Excel.Worksheet';
        default:
            return undefined;
    }
}

/** Translates an AST symbol's nameSpan to a Location in its owning module. */
function astSymbolToLocation(
    xlsmPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    symbol: AstSymbol,
): vscode.Location | undefined {
    const mod = byModule.get(symbol.moduleName.toLowerCase());
    if (!mod) { return undefined; }
    return new vscode.Location(
        encodeModuleUri(xlsmPath, mod.moduleName),
        new vscode.Range(
            offsetToPosition(mod.source, symbol.nameSpan.start),
            offsetToPosition(mod.source, symbol.nameSpan.end),
        ),
    );
}

function projectMemberDefinitionToLocation(
    xlsmPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    definition: VbaProjectClassMemberDefinition,
): vscode.Location | undefined {
    const mod = byModule.get(definition.moduleName.toLowerCase());
    if (!mod) { return undefined; }
    return new vscode.Location(
        encodeModuleUri(xlsmPath, mod.moduleName),
        new vscode.Range(
            offsetToPosition(mod.source, definition.nameSpan.start),
            offsetToPosition(mod.source, definition.nameSpan.end),
        ),
    );
}

function sourceMemberDefinitionsAt(
    source: string,
    memberName: string,
    memberEndOffset: number,
    project: ProjectIndex,
    modules: VbaModuleSymbols[],
    currentModuleName: string,
    currentModuleType?: string,
    currentDocumentType?: EventHandlerDocumentType,
): readonly VbaProjectClassMemberDefinition[] {
    return resolveMemberDefinitionsAt(source, memberEndOffset, memberName, {
        codeNames: codeNamesForModules(modules),
        meType: meHostTypeForModule(currentModuleName, currentModuleType, currentDocumentType),
        meProjectType: meProjectTypeForModule(currentModuleName, currentModuleType),
        projectClassMembers: project.projectMemberSurfaces(currentModuleName),
    });
}

function projectClassMemberAtDefinition(
    project: ProjectIndex,
    moduleName: string,
    memberName: string,
    offset: number,
): VbaProjectClassMember | undefined {
    for (const type of project.projectMemberSurfaces(moduleName)) {
        if (type.moduleName.toLowerCase() !== moduleName.toLowerCase()) {
            continue;
        }
        const member = type.members.find((candidate) =>
            candidate.name.toLowerCase() === memberName.toLowerCase() &&
            (candidate.definitions ?? []).some((definition) =>
                offset >= definition.nameSpan.start && offset <= definition.nameSpan.end,
            ),
        );
        if (member) { return member; }
    }
    return undefined;
}

function memberDefinitionKey(definition: VbaProjectClassMemberDefinition): string {
    return `${definition.moduleName.toLowerCase()}:${definition.nameSpan.start}`;
}

function projectMemberReferenceLocations(
    xlsmPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    project: ProjectIndex,
    modules: VbaModuleSymbols[],
    memberName: string,
    definitions: readonly VbaProjectClassMemberDefinition[],
    includeDeclaration: boolean,
): vscode.Location[] {
    const targetKeys = new Set(definitions.map(memberDefinitionKey));
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
            const loc = projectMemberDefinitionToLocation(xlsmPath, byModule, definition);
            if (loc) { push(loc); }
        }
    }

    const codeNames = codeNamesForModules(modules);
    for (const mod of byModule.values()) {
        const occurrences = findIdentifierOccurrences(mod.source, memberName);
        if (occurrences.length === 0) { continue; }
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        const ctx: MemberCompletionContext = {
            codeNames,
            meType: meHostTypeForModule(mod.moduleName, mod.type, mod.documentType),
            meProjectType: meProjectTypeForModule(mod.moduleName, mod.type),
            projectClassMembers: project.projectMemberSurfaces(mod.moduleName),
        };
        // Tokenize the module once and hand each occurrence its prefix slice;
        // occurrence offsets ascend, so one pass over the tokens suffices.
        const moduleTokens = tokenizeCached(mod.source).filter((t) => t.kind !== 'comment');
        let tokenEnd = 0;
        for (const occ of occurrences) {
            const occEnd = occ.offset + memberName.length;
            while (tokenEnd < moduleTokens.length && moduleTokens[tokenEnd].end <= occEnd) {
                tokenEnd += 1;
            }
            const resolved = resolveMemberDefinitionsAt(
                mod.source,
                occEnd,
                memberName,
                ctx,
                moduleTokens.slice(0, tokenEnd),
            );
            if (!resolved.some((definition) => targetKeys.has(memberDefinitionKey(definition)))) {
                continue;
            }
            push(new vscode.Location(
                uri,
                new vscode.Range(occ.line, occ.column, occ.line, occ.column + memberName.length),
            ));
        }
    }
    return out;
}

function projectMemberRenameLocations(
    xlsmPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    project: ProjectIndex,
    modules: VbaModuleSymbols[],
    memberName: string,
    definitions: readonly VbaProjectClassMemberDefinition[],
): vscode.Location[] {
    const out = projectMemberReferenceLocations(
        xlsmPath,
        byModule,
        project,
        modules,
        memberName,
        definitions,
        true,
    );
    const seen = new Set(out.map((loc) =>
        `${loc.uri.toString()}:${loc.range.start.line}:${loc.range.start.character}`,
    ));
    const push = (location: vscode.Location): void => {
        const key = `${location.uri.toString()}:${location.range.start.line}:${location.range.start.character}`;
        if (!seen.has(key)) {
            seen.add(key);
            out.push(location);
        }
    };

    for (const definition of definitions) {
        const mod = byModule.get(definition.moduleName.toLowerCase());
        if (!mod) { continue; }
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        for (const occ of findIdentifierOccurrences(mod.source, memberName)) {
            if (occ.offset < definition.fullSpan.start || occ.offset > definition.fullSpan.end) {
                continue;
            }
            push(new vscode.Location(
                uri,
                new vscode.Range(occ.line, occ.column, occ.line, occ.column + memberName.length),
            ));
        }
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

/**
 * Collects the in-scope textual occurrences of `word` for a resolved reference
 * scope, honoring local-procedure restriction and shadowing exclusions.
 */
function occurrencesInScope(
    xlsmPath: string,
    byModule: Map<string, VbaModuleSymbols>,
    scope: ReferenceScope,
    word: string,
): vscode.Location[] {
    const out: vscode.Location[] = [];
    for (const moduleName of scope.searchModules) {
        const mod = byModule.get(moduleName.toLowerCase());
        if (!mod) { continue; }
        const exclusions = scope.shadowedSpans
            .filter((s) => s.moduleName.toLowerCase() === moduleName.toLowerCase())
            .map((s) => s.span);
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        for (const occ of findIdentifierOccurrences(mod.source, word)) {
            if (scope.kind === 'local' && scope.procedureSpan) {
                if (occ.offset < scope.procedureSpan.start || occ.offset > scope.procedureSpan.end) {
                    continue;
                }
            }
            if (exclusions.some((sp) => occ.offset >= sp.start && occ.offset <= sp.end)) {
                continue;
            }
            out.push(new vscode.Location(
                uri,
                new vscode.Range(occ.line, occ.column, occ.line, occ.column + word.length),
            ));
        }
    }
    return out;
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
        const source = document.getText();
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
        const workbookPaths = new Set<string>();
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                workbookPaths.add(decodeModuleUri(document.uri).xlsmPath);
            } catch {
                // Ignore malformed XLIDE URIs.
            }
        }

        for (const xlsmPath of workbookPaths) {
            if (token.isCancellationRequested) { return out; }
            try {
                const context = await this._projectIndexService.contextForWorkbook(xlsmPath, 'live');
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
                        encodeModuleUri(xlsmPath, mod.moduleName),
                        toPosition,
                        symbol,
                    ));
                }
            } catch {
                // Workspace symbols are best-effort; skip workbooks that fail to read.
            }
        }

        for (const document of vscode.workspace.textDocuments) {
            if (token.isCancellationRequested) { return out; }
            if (!isStandaloneVbaDocument(document)) {
                continue;
            }
            const source = document.getText();
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
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (token?.isCancellationRequested) { return undefined; }

        const documentVersion = document.version;
        const source = document.getText();
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

        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const context = await this._projectIndexService.contextForWorkbook(xlsmPath, 'live');
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
                    projectMemberDefinitionToLocation(xlsmPath, byModule, definition),
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
                    .map((definition) => projectTypeDefinitionToLocation(xlsmPath, byModule, definition))
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
            const loc = astSymbolToLocation(xlsmPath, byModule, sym);
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
        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const navigation = await this._projectIndexService.contextForWorkbook(xlsmPath, 'live');
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
                xlsmPath,
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
        const targetMemberDefinitions = memberDefinitions.length > 0
            ? memberDefinitions
            : memberAtDefinition?.definitions ?? [];
        if (targetMemberDefinitions.length > 0) {
            return projectMemberReferenceLocations(
                xlsmPath,
                byModule,
                project,
                modules,
                word,
                targetMemberDefinitions,
                context.includeDeclaration,
            );
        }

        const typeReference = resolveTypeReferenceAt(
            source,
            document.offsetAt(position),
            { projectTypes: project.visibleTypeNames(moduleName) },
        );
        if (typeReference) {
            const definitions = typeDefinitionsForReference(project, moduleName, typeReference);
            if (definitions.length > 0) {
                return typeReferenceLocations(
                    xlsmPath,
                    byModule,
                    project,
                    typeReference.name,
                    definitions,
                    includeTypeDeclaration(context.includeDeclaration, definitions),
                );
            }
        }

        const scope = project.referenceScope(
            moduleName,
            word,
            document.offsetAt(position),
        );
        const locations = occurrencesInScope(xlsmPath, byModule, scope, word);

        if (context.includeDeclaration) {
            return locations;
        }
        // Drop the declaration identifier(s) themselves from the result set.
        const declKeys = new Set(
            scope.definitions.map(
                (d) => `${d.moduleName.toLowerCase()}:${d.nameSpan.start}`,
            ),
        );
        return locations.filter((loc) => {
            const mod = decodeModuleUri(loc.uri).moduleName.toLowerCase();
            const m = byModule.get(mod);
            if (!m) { return true; }
            const offset = (lineStartOffsets(m.source)[loc.range.start.line] ?? 0)
                + loc.range.start.character;
            return !declKeys.has(`${mod}:${offset}`);
        });
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

        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const navigation = await this._projectIndexService.contextForWorkbook(xlsmPath, 'strict');
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
            throw new Error(`'${word}' is not a renameable VBA symbol in this workbook.`);
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
        if (!VBA_IDENTIFIER_NAME_RE.test(newName)) {
            throw new Error(`'${newName}' is not a valid VBA identifier.`);
        }
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const oldName = document.getText(wordRange);
        if (oldName.toLowerCase() === newName.toLowerCase()) { return undefined; }

        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const navigation = await this._projectIndexService.contextForWorkbook(xlsmPath, 'strict');
        if (token?.isCancellationRequested || document.version !== documentVersion) {
            return undefined;
        }
        const { modules, project, byModule } = navigation;
        const current = byModule.get(moduleIdentityKey(moduleName));

        const memberDefinitions = sourceMemberDefinitionsAt(
            source,
            oldName,
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
                oldName,
                document.offsetAt(position),
            );
        const targetMemberDefinitions = memberDefinitions.length > 0
            ? memberDefinitions
            : memberAtDefinition?.definitions ?? [];
        if (targetMemberDefinitions.length > 0) {
            const edit = new vscode.WorkspaceEdit();
            for (const loc of projectMemberRenameLocations(
                xlsmPath,
                byModule,
                project,
                modules,
                oldName,
                targetMemberDefinitions,
            )) {
                edit.replace(loc.uri, loc.range, newName);
            }
            return edit;
        }

        const scope = project.referenceScope(
            moduleName,
            oldName,
            document.offsetAt(position),
        );
        if (scope.definitions.length === 0) {
            throw new Error(`'${oldName}' is not a renameable VBA symbol in this workbook.`);
        }

        const edit = new vscode.WorkspaceEdit();
        for (const loc of occurrencesInScope(xlsmPath, byModule, scope, oldName)) {
            edit.replace(loc.uri, loc.range, newName);
        }
        return edit;
    }
}
