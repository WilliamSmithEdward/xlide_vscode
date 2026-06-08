import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import {
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
    encodeModuleUri,
    moduleIdentityKey,
} from './xlideFileSystem';
import { VbaSymbolIndex, VbaModuleSymbols, parseVbaModule } from './vbaSymbolIndex';
import { applyOpenDocumentSources } from './vbaOpenDocuments';
import {
    analyzeVbaStructure,
    stripVba,
    detectSmartBlockOpener,
    findIdentifierOccurrences,
    isSmartBlockClosedAhead,
    lineStartOffsets,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    smartBlockInsertion,
    VBA_IDENTIFIER_NAME_RE,
    VBA_IDENTIFIER_RE,
    withMemberContinuationText,
} from './vbaStructuralAnalysis';
import {
    diagnosticSourceForCode,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    isXlideDiagnosticSource,
    normalizeDiagnosticCode,
    ProjectIndex,
    ReferenceScope,
    resolveDiagnosticCodeActions,
    resolveMemberCompletions,
    resolveProcedureLabelDefinitionAt,
    resolveTypeReferenceAt,
    resolveTypeSemanticTokens,
    TypeSemanticTokenType,
    type VbaDiagnosticData,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    type ModuleSymbolKind,
    type Span,
    VbaSymbol as AstSymbol,
    type VbaSymbolKind,
} from './analyzer';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import { registerVbaMemberCompletion } from './vbaMemberCompletion';
import { DocMetadataLoader } from './vbaDocMetadata';
import {
    buildVbaProjectIndex as buildProjectIndex,
    moduleKindFromType,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeDefinitionsForReference,
    typeReferenceLocations,
} from './vbaNavigation';
import {
    buildLiveVbaProjectIndex,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import {
    documentOutlineSymbolsForSource,
    workspaceSymbols as presentedWorkspaceSymbols,
    type VbaPresentedSymbol,
    type VbaPresentedWorkspaceSymbol,
} from './vbaSymbolPresentation';
import {
    isAnalysisRuleTracked,
} from './analysisSettingsCore';
import { effectiveWorkbookAnalysisSettings } from './workbookAnalysisSettings';
import { isWorkbookSettingsError, settingsPathForWorkbook } from './workbookSettings';
import {
    validateXlideGlobalSettingsFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    xlideEditorBlockLayoutFromConfig,
    type XlideGlobalSettingsProblem,
} from './globalSettings';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME, language: XLIDE_VBA_LANGUAGE_ID },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
    { language: XLIDE_VBA_LANGUAGE_ID },
];
const XLIDE_SOURCE_ACTION_KIND = vscode.CodeActionKind.Source.append('xlide');
const XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('analyzeCurrentModule');
const XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('exportCurrentModule');
const XLIDE_DIAGNOSTIC_DATA = Symbol('xlideDiagnosticData');

type XlideDiagnosticWithData = vscode.Diagnostic & {
    [XLIDE_DIAGNOSTIC_DATA]?: VbaDiagnosticData;
};

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

function spanRange(source: string, span: Span): vscode.Range {
    return new vscode.Range(
        offsetToPosition(source, span.start),
        offsetToPosition(source, span.end),
    );
}

function documentSymbolToVscode(source: string, symbol: VbaPresentedSymbol): vscode.DocumentSymbol {
    const out = new vscode.DocumentSymbol(
        symbol.name,
        symbol.detail,
        astSymbolKindToVscode(symbol.kind),
        spanRange(source, symbol.fullSpan),
        spanRange(source, symbol.nameSpan),
    );
    out.children = symbol.children.map((child) => documentSymbolToVscode(source, child));
    return out;
}

function workspaceSymbolToVscode(
    uri: vscode.Uri,
    source: string,
    symbol: VbaPresentedWorkspaceSymbol,
): vscode.SymbolInformation {
    return new vscode.SymbolInformation(
        symbol.name,
        astSymbolKindToVscode(symbol.kind),
        symbol.containerName,
        new vscode.Location(uri, spanRange(source, symbol.nameSpan)),
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

function moduleMapWithLiveDocument(
    modules: VbaModuleSymbols[],
    moduleName: string,
    source: string,
    type?: string,
    documentType?: EventHandlerDocumentType,
): Map<string, VbaModuleSymbols> {
    const byModule = new Map(modules.map((m) => [
        moduleIdentityKey(m.moduleName),
        { ...m, symbols: parseVbaModule(m.source) },
    ]));
    byModule.set(moduleIdentityKey(moduleName), {
        moduleName,
        type,
        documentType,
        source,
        symbols: parseVbaModule(source),
    });
    return byModule;
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
    const member = resolveMemberCompletions(source, memberEndOffset, {
        codeNames: codeNamesForModules(modules),
        meType: meHostTypeForModule(currentModuleName, currentModuleType, currentDocumentType),
        meProjectType: meProjectTypeForModule(currentModuleName, currentModuleType),
        projectClassMembers: project.projectMemberSurfaces(currentModuleName),
    }).find((item) => item.name.toLowerCase() === memberName.toLowerCase());
    return member?.definitions ?? [];
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

    for (const mod of byModule.values()) {
        const uri = encodeModuleUri(xlsmPath, mod.moduleName);
        for (const occ of findIdentifierOccurrences(mod.source, memberName)) {
            const resolved = sourceMemberDefinitionsAt(
                mod.source,
                memberName,
                occ.offset + memberName.length,
                project,
                modules,
                mod.moduleName,
                mod.type,
                mod.documentType,
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

class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    provideDocumentSymbols(
        document: vscode.TextDocument,
        token: vscode.CancellationToken,
    ): vscode.DocumentSymbol[] {
        if (!isVbaDocument(document)) { return []; }
        if (token.isCancellationRequested) { return []; }
        const source = document.getText();
        const moduleName = moduleNameFromDocument(document);
        return documentOutlineSymbolsForSource(
            moduleName,
            moduleKindFromDocument(document),
            source,
        ).map((symbol) => documentSymbolToVscode(source, symbol));
    }
}

class VbaWorkspaceSymbolProvider implements vscode.WorkspaceSymbolProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

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
                const modules = applyOpenDocumentSources(
                    await this._index.getAllModules(xlsmPath),
                    xlsmPath,
                );
                const project = buildLiveVbaProjectIndex(modules);
                const byModule = new Map(modules.map((mod) => [
                    mod.moduleName.toLowerCase(),
                    mod,
                ]));
                for (const symbol of presentedWorkspaceSymbols(project, query)) {
                    const mod = byModule.get(symbol.moduleName.toLowerCase());
                    if (!mod) { continue; }
                    out.push(workspaceSymbolToVscode(
                        encodeModuleUri(xlsmPath, mod.moduleName),
                        mod.source,
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
            const project = buildLiveVbaProjectIndex([], {
                moduleName,
                moduleKind: moduleKindFromDocument(document),
                source,
            });
            for (const symbol of presentedWorkspaceSymbols(project, query)) {
                out.push(workspaceSymbolToVscode(document.uri, source, symbol));
            }
        }

        return out;
    }
}

class VbaDefinitionProvider implements vscode.DefinitionProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideDefinition(
        document: vscode.TextDocument,
        position: vscode.Position,
    ): Promise<vscode.Location[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }

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
        const modules = applyOpenDocumentSources(
            await this._index.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const project = buildLiveVbaProjectIndex(modules, {
            moduleName,
            moduleKind: moduleKindFromType(current?.type),
            source,
        });
        const byModule = moduleMapWithLiveDocument(
            modules,
            moduleName,
            source,
            current?.type,
            current?.documentType,
        );

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

class VbaReferenceProvider implements vscode.ReferenceProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideReferences(
        document: vscode.TextDocument,
        position: vscode.Position,
        context: vscode.ReferenceContext,
    ): Promise<vscode.Location[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const modules = applyOpenDocumentSources(
            await this._index.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const project = buildLiveVbaProjectIndex(modules, {
            moduleName,
            moduleKind: moduleKindFromType(current?.type),
            source,
        });
        const byModule = moduleMapWithLiveDocument(
            modules,
            moduleName,
            source,
            current?.type,
            current?.documentType,
        );

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

class VbaRenameProvider implements vscode.RenameProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async prepareRename(
        document: vscode.TextDocument,
        position: vscode.Position,
    ): Promise<vscode.Range | { range: vscode.Range; placeholder: string }> {
        if (document.uri.scheme !== XLIDE_SCHEME) {
            throw new Error('Rename is only supported in XLIDE VBA modules.');
        }
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { throw new Error('No symbol at cursor.'); }
        const word = document.getText(wordRange);

        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const modules = applyOpenDocumentSources(
            await this._index.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const project = buildProjectIndex(modules, {
            moduleName,
            moduleKind: moduleKindFromType(current?.type),
            source,
        });

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
    ): Promise<vscode.WorkspaceEdit | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (!VBA_IDENTIFIER_NAME_RE.test(newName)) {
            throw new Error(`'${newName}' is not a valid VBA identifier.`);
        }
        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const oldName = document.getText(wordRange);
        if (oldName.toLowerCase() === newName.toLowerCase()) { return undefined; }

        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const modules = applyOpenDocumentSources(
            await this._index.getAllModules(xlsmPath),
            xlsmPath,
        );
        const current = modules.find(
            (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        const project = buildProjectIndex(modules, {
            moduleName,
            moduleKind: moduleKindFromType(current?.type),
            source,
        });
        const byModule = moduleMapWithLiveDocument(
            modules,
            moduleName,
            source,
            current?.type,
            current?.documentType,
        );

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

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

function isVbaDocument(document: vscode.TextDocument): boolean {
    return document.languageId === 'vba'
        || document.languageId === XLIDE_VBA_LANGUAGE_ID
        || document.uri.scheme === XLIDE_SCHEME;
}

function isStandaloneVbaDocument(document: vscode.TextDocument): boolean {
    return document.uri.scheme !== XLIDE_SCHEME && document.languageId === 'vba';
}

function moduleNameFromDocument(document: vscode.TextDocument): string {
    if (document.uri.scheme === XLIDE_SCHEME) {
        try {
            return decodeModuleUri(document.uri).moduleName;
        } catch {
            /* fall through */
        }
    }
    const base = document.uri.path.split('/').pop() ?? 'Module';
    return base.replace(/\.[^.]+$/, '') || 'Module';
}

function moduleKindFromDocument(document: vscode.TextDocument): ModuleSymbolKind {
    const fileName = document.uri.path.split('/').pop() ?? '';
    if (/\.cls$/i.test(fileName)) {
        return 'class';
    }
    if (/\.frm$/i.test(fileName)) {
        return 'userform';
    }
    return 'standard';
}

async function liveProjectIndexForDocument(
    index: VbaSymbolIndex,
    document: vscode.TextDocument,
    source: string,
    moduleName: string,
): Promise<ProjectIndex> {
    if (document.uri.scheme !== XLIDE_SCHEME) {
        return buildLiveVbaProjectIndex([], {
            moduleName,
            moduleKind: moduleKindFromDocument(document),
            source,
        });
    }

    const decoded = decodeModuleUri(document.uri);
    const modules = applyOpenDocumentSources(
        await index.getAllModules(decoded.xlsmPath),
        decoded.xlsmPath,
    );
    const current = modules.find(
        (m) => m.moduleName.toLowerCase() === moduleName.toLowerCase(),
    );
    return buildLiveVbaProjectIndex(modules, {
        moduleName,
        moduleKind: moduleKindFromType(current?.type),
        source,
    });
}

const TYPE_TOKEN_TYPES: TypeSemanticTokenType[] = [
    'class',
    'enum',
    'struct',
    'type',
];
const TYPE_TOKEN_LEGEND = new vscode.SemanticTokensLegend(TYPE_TOKEN_TYPES);
const TYPE_SEMANTIC_PROJECT_TYPES_CACHE_TTL_MS = 5000;
const DIAGNOSTIC_OPEN_LOCAL_DELAY_MS = 25;
const DIAGNOSTIC_OPEN_FULL_DELAY_MS = 150;
const DIAGNOSTIC_EDIT_LOCAL_DELAY_MS = 90;
const DIAGNOSTIC_EDIT_FULL_DELAY_MS = 450;

interface CachedTypeSemanticProjectTypes {
    at: number;
    documentVersion: number;
    projectTypes: VbaProjectAnalysisOptions['projectTypes'];
}

class VbaTypeSemanticTokensProvider implements vscode.DocumentSemanticTokensProvider {
    private readonly _onDidChangeSemanticTokens = new vscode.EventEmitter<void>();
    private readonly _projectTypesCache = new Map<string, CachedTypeSemanticProjectTypes>();
    private readonly _projectTypeRefreshes = new Set<string>();

    readonly onDidChangeSemanticTokens = this._onDidChangeSemanticTokens.event;

    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideDocumentSemanticTokens(
        document: vscode.TextDocument,
        token: vscode.CancellationToken,
    ): Promise<vscode.SemanticTokens> {
        const builder = new vscode.SemanticTokensBuilder(TYPE_TOKEN_LEGEND);
        if (!isVbaDocument(document)) { return builder.build(); }

        const source = document.getText();
        const moduleName = moduleNameFromDocument(document);
        const projectTypes = document.uri.scheme === XLIDE_SCHEME
            ? this._cachedProjectTypesForDocument(document, { requireFresh: false }) ?? []
            : await this._projectTypesForDocument(document, source, moduleName);
        if (
            document.uri.scheme === XLIDE_SCHEME &&
            !this._cachedProjectTypesForDocument(document, { requireFresh: true })
        ) {
            this._refreshProjectTypesInBackground(document, source, moduleName);
        }

        for (const item of resolveTypeSemanticTokens(source, { projectTypes })) {
            if (token.isCancellationRequested) { break; }
            builder.push(
                new vscode.Range(
                    document.positionAt(item.span.start),
                    document.positionAt(item.span.end),
                ),
                item.tokenType,
                [],
            );
        }
        return builder.build();
    }

    private _cachedProjectTypesForDocument(
        document: vscode.TextDocument,
        options: { requireFresh?: boolean } = {},
    ): VbaProjectAnalysisOptions['projectTypes'] | undefined {
        const cached = this._projectTypesCache.get(document.uri.toString());
        if (!cached || cached.documentVersion !== document.version) {
            return undefined;
        }
        if (
            options.requireFresh &&
            Date.now() - cached.at >= TYPE_SEMANTIC_PROJECT_TYPES_CACHE_TTL_MS
        ) {
            return undefined;
        }
        return cached.projectTypes;
    }

    private _refreshProjectTypesInBackground(
        document: vscode.TextDocument,
        source: string,
        moduleName: string,
    ): void {
        const key = document.uri.toString();
        if (this._projectTypeRefreshes.has(key)) {
            return;
        }
        this._projectTypeRefreshes.add(key);
        void this._projectTypesForDocument(document, source, moduleName)
            .then(() => this._onDidChangeSemanticTokens.fire())
            .finally(() => this._projectTypeRefreshes.delete(key));
    }

    private async _projectTypesForDocument(
        document: vscode.TextDocument,
        source: string,
        moduleName: string,
    ): Promise<VbaProjectAnalysisOptions['projectTypes']> {
        const key = document.uri.toString();
        const cached = this._cachedProjectTypesForDocument(document, { requireFresh: true });
        if (cached) {
            return cached;
        }

        const previous = this._projectTypesCache.get(key);
        const documentVersion = document.version;

        try {
            const project = await liveProjectIndexForDocument(
                this._index,
                document,
                source,
                moduleName,
            );
            const projectTypes = projectAnalysisOptionsForModule(project, moduleName).projectTypes ?? [];
            this._projectTypesCache.set(key, { at: Date.now(), documentVersion, projectTypes });
            return projectTypes;
        } catch {
            return previous?.projectTypes ?? [];
        }
    }
}

/**
 * Live diagnostics: structural block-balance (analyzeVbaStructure) plus the analyzer's
 * high-confidence semantic rules (analyzeModule) - unterminated strings,
 * duplicate procedures/declarations, assignment to a constant, and a
 * configurable Option Explicit reminder. Runs on open and (debounced) on every
 * edit so problems surface while typing, the way a real IDE does. No save and
 * no Python round-trip required - everything is computed from the editor text.
 */
function registerVbaDiagnostics(
    context: vscode.ExtensionContext,
    index: VbaSymbolIndex,
): void {
    const collection = vscode.languages.createDiagnosticCollection('vba');
    const localTimers = new Map<string, ReturnType<typeof setTimeout>>();
    const fullTimers = new Map<string, ReturnType<typeof setTimeout>>();
    const diagnosticGenerations = new Map<string, number>();
    const completedFullGenerations = new Map<string, number>();
    const workbookSettingsWatchers = new Map<string, vscode.Disposable[]>();

    type DiagnosticPassKind = 'local' | 'full';

    interface DiagnosticScheduleDelays {
        localDelayMs: number;
        fullDelayMs?: number;
    }

    const severityToVscode = (s: RuleSeverity): vscode.DiagnosticSeverity => {
        switch (s) {
            case 'error': return vscode.DiagnosticSeverity.Error;
            case 'warning': return vscode.DiagnosticSeverity.Warning;
            case 'information': return vscode.DiagnosticSeverity.Information;
        }
    };

    const run = (
        document: vscode.TextDocument,
        delays: DiagnosticScheduleDelays = { localDelayMs: 0, fullDelayMs: 0 },
    ): void => {
        schedule(document, delays);
    };

    const runPass = (
        document: vscode.TextDocument,
        generation: number,
        pass: DiagnosticPassKind,
    ): void => {
        void runPassAsync(document, generation, pass).catch((err) => {
            if (!isVbaDocument(document)) {
                return;
            }
            const key = document.uri.toString();
            if (!isCurrentDiagnosticRun(document, key, generation, document.version)) {
                return;
            }
            collection.set(document.uri, [diagnosticForAnalysisRunError(document, err)]);
        });
    };

    const diagnosticForAnalysisRunError = (
        document: vscode.TextDocument,
        err: unknown,
    ): vscode.Diagnostic => {
        const firstLine = document.lineCount > 0 ? document.lineAt(0).text : '';
        const range = new vscode.Range(0, 0, 0, Math.min(firstLine.length, 1));
        const settingsError = isWorkbookSettingsError(err);
        const message = settingsError
            ? `${err.message} Fix or delete the workbook settings sidecar.`
            : `XLIDE diagnostics failed: ${err instanceof Error ? err.message : String(err)}`;
        const diagnostic = new vscode.Diagnostic(
            range,
            message,
            vscode.DiagnosticSeverity.Error,
        );
        diagnostic.source = settingsError ? 'XLIDE/settings' : 'XLIDE';
        diagnostic.code = settingsError ? 'workbook-settings-invalid' : 'diagnostics-failed';
        return diagnostic;
    };

    const diagnosticsForGlobalSettingsProblems = (
        document: vscode.TextDocument,
        problems: readonly XlideGlobalSettingsProblem[],
    ): vscode.Diagnostic[] => {
        const firstLine = document.lineCount > 0 ? document.lineAt(0).text : '';
        const range = new vscode.Range(0, 0, 0, Math.min(firstLine.length, 1));
        return problems.map((problem) => {
            const diagnostic = new vscode.Diagnostic(
                range,
                `${problem.message} Fix the value in VS Code settings.`,
                vscode.DiagnosticSeverity.Error,
            );
            diagnostic.source = 'XLIDE/settings';
            diagnostic.code = 'global-setting-invalid';
            return diagnostic;
        });
    };

    const runPassAsync = async (
        document: vscode.TextDocument,
        generation: number,
        pass: DiagnosticPassKind,
    ): Promise<void> => {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const documentVersion = document.version;
        const config = vscode.workspace.getConfiguration('xlide');
        const settingsDiagnostics = diagnosticsForGlobalSettingsProblems(
            document,
            validateXlideGlobalSettingsFromConfig(config),
        );
        if (!xlideDiagnosticsEnabledFromConfig(config).value) {
            if (!isCurrentDiagnosticRun(document, key, generation, documentVersion)) {
                return;
            }
            if (settingsDiagnostics.length > 0) {
                collection.set(document.uri, settingsDiagnostics);
            } else {
                collection.delete(document.uri);
            }
            return;
        }
        const text = document.getText();

        const moduleName = moduleNameFromDocument(document);
        let workbookPath: string | undefined;
        let moduleType: string | undefined;
        let moduleKind: ModuleSymbolKind | undefined;
        let documentType: EventHandlerDocumentType | undefined;
        let projectOptions: VbaProjectAnalysisOptions = {};
        if (document.uri.scheme === XLIDE_SCHEME) {
            try {
                const { xlsmPath } = decodeModuleUri(document.uri);
                workbookPath = xlsmPath;
                ensureWorkbookSettingsWatcher(xlsmPath);
                if (pass === 'full') {
                    const modules = applyOpenDocumentSources(
                        await index.getAllModules(xlsmPath),
                        xlsmPath,
                    );
                    const current = modules.find(
                        (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
                    );
                    moduleType = current?.type;
                    moduleKind = moduleKindFromType(current?.type);
                    documentType = current?.documentType;
                    const project = buildLiveVbaProjectIndex(modules, {
                        moduleName,
                        moduleKind,
                        source: text,
                    });
                    projectOptions = projectAnalysisOptionsForModule(
                        project,
                        moduleName,
                        projectProcedureSignatures(project),
                    );
                }
            } catch {
                projectOptions = {};
            }
        }

        const analysisSettings = await effectiveWorkbookAnalysisSettings(workbookPath);
        const activeEditor = vscode.window.activeTextEditor;
        const activeIncompleteExpressionOffset = activeEditor?.document === document
            ? document.offsetAt(activeEditor.selection.active)
            : undefined;
        const moduleAnalysis = analyzeVbaModuleSource({
            source: text,
            moduleName,
            moduleType,
            moduleKind,
            documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
            activeIncompleteExpressionOffset,
        });
        const diagnostics = diagnosticsFromModuleAnalysis(
            document,
            moduleAnalysis,
            analysisSettings.untrackedRules,
            settingsDiagnostics,
        );
        publishDiagnosticsIfCurrent(document, key, generation, documentVersion, pass, diagnostics);
    };

    const diagnosticsFromModuleAnalysis = (
        document: vscode.TextDocument,
        moduleAnalysis: ReturnType<typeof analyzeVbaModuleSource>,
        untrackedRules: readonly string[],
        settingsDiagnostics: readonly vscode.Diagnostic[],
    ): vscode.Diagnostic[] => {
        const diagnostics: vscode.Diagnostic[] = [...settingsDiagnostics];
        for (const d of moduleAnalysis.diagnostics) {
            if (!isAnalysisRuleTracked(d.code, untrackedRules)) {
                continue;
            }
            const diag = new vscode.Diagnostic(
                new vscode.Range(
                    document.positionAt(d.span.start),
                    document.positionAt(d.span.end),
                ),
                d.message,
                severityToVscode(d.severity),
            );
            diag.source = diagnosticSourceForCode(d.code);
            if (d.code) {
                diag.code = d.code;
            }
            if (d.data) {
                (diag as XlideDiagnosticWithData)[XLIDE_DIAGNOSTIC_DATA] = d.data;
            }
            diagnostics.push(diag);
        }
        return diagnostics;
    };

    const publishDiagnosticsIfCurrent = (
        document: vscode.TextDocument,
        key: string,
        generation: number,
        documentVersion: number,
        pass: DiagnosticPassKind,
        diagnostics: vscode.Diagnostic[],
    ): void => {
        if (!isCurrentDiagnosticRun(document, key, generation, documentVersion)) {
            return;
        }
        if (pass === 'local' && completedFullGenerations.get(key) === generation) {
            return;
        }
        if (pass === 'full') {
            completedFullGenerations.set(key, generation);
        }
        collection.set(document.uri, diagnostics);
    };

    const isCurrentDiagnosticRun = (
        document: vscode.TextDocument,
        key: string,
        generation: number,
        documentVersion: number,
    ): boolean => {
        return diagnosticGenerations.get(key) === generation &&
            document.version === documentVersion &&
            vscode.workspace.textDocuments.includes(document);
    };

    const nextDiagnosticGeneration = (key: string): number => {
        const next = (diagnosticGenerations.get(key) ?? 0) + 1;
        diagnosticGenerations.set(key, next);
        completedFullGenerations.delete(key);
        return next;
    };

    const clearTimer = (
        timers: Map<string, ReturnType<typeof setTimeout>>,
        key: string,
    ): void => {
        const existing = timers.get(key);
        if (existing) {
            clearTimeout(existing);
            timers.delete(key);
        }
    };

    const schedule = (
        document: vscode.TextDocument,
        delays: DiagnosticScheduleDelays = {
            localDelayMs: DIAGNOSTIC_EDIT_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_EDIT_FULL_DELAY_MS,
        },
    ): void => {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const generation = nextDiagnosticGeneration(key);
        clearTimer(localTimers, key);
        clearTimer(fullTimers, key);
        localTimers.set(key, setTimeout(() => {
            localTimers.delete(key);
            runPass(document, generation, 'local');
        }, delays.localDelayMs));
        if (document.uri.scheme === XLIDE_SCHEME && delays.fullDelayMs !== undefined) {
            fullTimers.set(key, setTimeout(() => {
                fullTimers.delete(key);
                runPass(document, generation, 'full');
            }, delays.fullDelayMs));
        }
    };

    const workbookKey = (workbookPath: string): string => path.resolve(workbookPath).toLowerCase();
    const rerunWorkbookDocuments = (workbookPath: string): void => {
        const key = workbookKey(workbookPath);
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                if (workbookKey(decodeModuleUri(document.uri).xlsmPath) === key) {
                    run(document);
                }
            } catch {
                // Ignore URIs that are no longer valid XLIDE module documents.
            }
        }
    };
    const ensureWorkbookSettingsWatcher = (workbookPath: string): void => {
        const key = workbookKey(workbookPath);
        if (workbookSettingsWatchers.has(key)) {
            return;
        }
        const settingsPath = settingsPathForWorkbook(workbookPath);
        const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
            path.dirname(settingsPath),
            path.basename(settingsPath),
        ));
        const rerun = () => rerunWorkbookDocuments(workbookPath);
        workbookSettingsWatchers.set(key, [
            watcher.onDidCreate(rerun),
            watcher.onDidChange(rerun),
            watcher.onDidDelete(rerun),
            watcher,
        ]);
    };
    const pruneWorkbookSettingsWatchers = (): void => {
        const openWorkbookKeys = new Set<string>();
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                openWorkbookKeys.add(workbookKey(decodeModuleUri(document.uri).xlsmPath));
            } catch {
                // Ignore invalid XLIDE URIs.
            }
        }
        for (const [key, disposables] of workbookSettingsWatchers) {
            if (openWorkbookKeys.has(key)) {
                continue;
            }
            disposables.forEach((disposable) => disposable.dispose());
            workbookSettingsWatchers.delete(key);
        }
    };
    const disposeWorkbookSettingsWatchers = (): void => {
        for (const disposables of workbookSettingsWatchers.values()) {
            disposables.forEach((disposable) => disposable.dispose());
        }
        workbookSettingsWatchers.clear();
    };

    context.subscriptions.push(
        collection,
        vscode.workspace.onDidOpenTextDocument((document) => schedule(document, {
            localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
        })),
        vscode.workspace.onDidChangeTextDocument((e) => schedule(e.document)),
        vscode.window.onDidChangeActiveTextEditor(() => {
            const editor = vscode.window.activeTextEditor;
            if (editor) {
                schedule(editor.document, {
                    localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
                    fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
                });
            }
        }),
        vscode.workspace.onDidCloseTextDocument((doc) => {
            const key = doc.uri.toString();
            clearTimer(localTimers, key);
            clearTimer(fullTimers, key);
            nextDiagnosticGeneration(key);
            collection.delete(doc.uri);
            pruneWorkbookSettingsWatchers();
        }),
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('xlide.diagnostics') ||
                e.affectsConfiguration('xlide.analysis')) {
                vscode.workspace.textDocuments.forEach((document) => run(document));
            }
        }),
        { dispose: disposeWorkbookSettingsWatchers },
    );
    vscode.workspace.textDocuments.forEach((document) => schedule(document, {
        localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
        fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
    }));
}

class VbaCodeActionProvider implements vscode.CodeActionProvider {
    public provideCodeActions(
        document: vscode.TextDocument,
        _range: vscode.Range,
        context: vscode.CodeActionContext,
        _token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.CodeAction[]> {
        if (!isVbaDocument(document)) { return []; }
        const wantsQuickFix = codeActionKindRequested(context.only, vscode.CodeActionKind.QuickFix);
        const wantsAnalyzeCurrentModule = codeActionKindRequested(context.only, XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND);
        const wantsExportCurrentModule = codeActionKindRequested(context.only, XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND);
        if (!wantsQuickFix && !wantsAnalyzeCurrentModule && !wantsExportCurrentModule) {
            return [];
        }

        const source = document.getText();
        const actions: vscode.CodeAction[] = [];
        if (wantsAnalyzeCurrentModule && document.uri.scheme === XLIDE_SCHEME) {
            const action = new vscode.CodeAction(
                'XLIDE: Analyze Current Module',
                XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.analyzeCurrentModule',
                title: 'Analyze Current Module',
            };
            actions.push(action);
        }
        if (wantsExportCurrentModule && document.uri.scheme === XLIDE_SCHEME && !document.uri.authority) {
            const action = new vscode.CodeAction(
                'XLIDE: Export/Sync Current Module',
                XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.exportCurrentModuleToFolder',
                title: 'Export/Sync Current Module',
            };
            actions.push(action);
        }
        if (!wantsQuickFix) {
            return actions;
        }
        let structuralDiagnostics: ReturnType<typeof analyzeVbaStructure> | undefined;
        for (const diagnostic of context.diagnostics) {
            if (!isXlideDiagnosticSource(diagnostic.source)) { continue; }
            const code = normalizeDiagnosticCode(diagnostic.code);
            if (!code) { continue; }
            const structuralDiagnostic = code === 'missing-block-closer'
                ? matchingStructuralDiagnostic(
                    diagnostic,
                    structuralDiagnostics ??= analyzeVbaStructure(source),
                )
                : undefined;
            const fixes = resolveDiagnosticCodeActions(source, {
                code,
                message: diagnostic.message,
                expectedClose: structuralDiagnostic?.expectedClose,
                insertLine: structuralDiagnostic?.insertLine,
                expectedCloseReplacementSpan: structuralDiagnostic?.expectedCloseReplacement
                    ? {
                        start: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.startCol,
                        )),
                        end: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.endCol,
                        )),
                    }
                    : undefined,
                expectedCloseReplacementText: structuralDiagnostic?.expectedCloseReplacement?.text,
                span: {
                    start: document.offsetAt(diagnostic.range.start),
                    end: document.offsetAt(diagnostic.range.end),
                },
                includeSuppressionAction: true,
                data: (diagnostic as XlideDiagnosticWithData)[XLIDE_DIAGNOSTIC_DATA],
            });
            for (const fix of fixes) {
                const action = new vscode.CodeAction(fix.title, vscode.CodeActionKind.QuickFix);
                action.diagnostics = [diagnostic];
                action.isPreferred = fix.isPreferred;
                const edit = new vscode.WorkspaceEdit();
                for (const textEdit of fix.edits) {
                    edit.replace(
                        document.uri,
                        new vscode.Range(
                            document.positionAt(textEdit.span.start),
                            document.positionAt(textEdit.span.end),
                        ),
                        textEdit.newText,
                    );
                }
                action.edit = edit;
                actions.push(action);
            }
        }
        return actions;
    }
}

function codeActionKindRequested(
    only: vscode.CodeActionKind | undefined,
    kind: vscode.CodeActionKind,
): boolean {
    return !only || only.contains(kind) || kind.contains(only);
}

function matchingStructuralDiagnostic(
    diagnostic: vscode.Diagnostic,
    problems: ReturnType<typeof analyzeVbaStructure>,
): ReturnType<typeof analyzeVbaStructure>[number] | undefined {
    return problems.find((problem) =>
        problem.code === normalizeDiagnosticCode(diagnostic.code) &&
        problem.message === diagnostic.message &&
        diagnostic.range.start.isEqual(new vscode.Position(problem.line, problem.startCol)) &&
        diagnostic.range.end.isEqual(new vscode.Position(problem.line, problem.endCol)),
    );
}

/**
 * VBA-IDE-style smart Enter: typing a block opener and pressing Enter
 * auto-inserts the matching closer below, leaving the cursor on the indented
 * body line. `With` also seeds the body line with `.` so member completion can
 * start immediately.
 */
function registerVbaAutoBlock(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        // React only to a plain Enter (newline plus optional auto-indent),
        // never to pastes or multi-character insertions.
        if (!/^\r?\n[ \t]*$/.test(change.text)) { return; }

        const openerLineIndex = change.range.start.line;
        const openerLine = doc.lineAt(openerLineIndex).text;
        const headerParensEdit = procedureHeaderParensEdit(openerLine);
        const normalizedOpenerLine = headerParensEdit
            ? `${openerLine.slice(0, headerParensEdit.startCol)}${headerParensEdit.newText}${openerLine.slice(headerParensEdit.endCol)}`
            : openerLine;
        const opener = detectSmartBlockOpener(stripVba(normalizedOpenerLine));
        if (!opener) {
            await maybeContinueWithMemberLine(doc, openerLineIndex);
            return;
        }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = doc.getText().split(/\r\n|\r|\n/).map(stripVba);
        strippedLines[openerLineIndex] = stripVba(normalizedOpenerLine);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const bodyLine = doc.lineAt(bodyLineIndex).text;
        if (!/^[ \t]*$/.test(bodyLine)) { return; }
        const smartBlock = smartBlockInsertion(normalizedOpenerLine, bodyLine, opener, {
            eol,
            insertCloser: !closedAhead,
            layout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
        });
        const bodyRange = new vscode.Range(
            new vscode.Position(bodyLineIndex, 0),
            new vscode.Position(bodyLineIndex, bodyLine.length),
        );

        applying = true;
        try {
            await editor.edit(
                (eb) => {
                    if (headerParensEdit) {
                        eb.insert(
                            new vscode.Position(openerLineIndex, headerParensEdit.startCol),
                            headerParensEdit.newText,
                        );
                    }
                    eb.replace(
                        bodyRange,
                        smartBlock.replacementText,
                    );
                },
                { undoStopBefore: false, undoStopAfter: true },
            );
        } finally {
            applying = false;
        }

        // Keep the caret on the indented body line, above the inserted End. The
        // delayed pass wins same-Enter listener races such as canonical casing.
        const placeCaret = (): void => {
            if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
                return;
            }
            const caretLineIndex = bodyLineIndex + smartBlock.bodyLineOffset;
            if (caretLineIndex >= doc.lineCount || doc.lineAt(caretLineIndex).text !== smartBlock.bodyText) {
                return;
            }
            const caret = new vscode.Position(
                caretLineIndex,
                smartBlock.bodyText.length,
            );
            editor.selection = new vscode.Selection(caret, caret);
        };
        placeCaret();
        setTimeout(placeCaret, 0);
    });

    context.subscriptions.push(sub);
}

async function maybeContinueWithMemberLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<void> {
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return; }

    const lineText = withMemberContinuationText(doc.getText(), previousLineIndex);
    if (!lineText) { return; }

    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return; }

    const placeCaret = (): void => {
        if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
            return;
        }
        if (bodyLineIndex >= doc.lineCount || doc.lineAt(bodyLineIndex).text !== lineText) {
            return;
        }
        const caret = new vscode.Position(bodyLineIndex, lineText.length);
        editor.selection = new vscode.Selection(caret, caret);
    };
    placeCaret();
    setTimeout(placeCaret, 0);
}

/**
 * Keeps simple loop iterator names paired across `For`/`For Each` and `Next`.
 * This intentionally lives outside snippets so hand-written loops get the same
 * behavior as completed loops.
 */
function registerVbaLoopIteratorSync(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        if (/[\r\n]/.test(change.text)) { return; }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const lineIndex = Math.min(change.range.start.line, doc.lineCount - 1);
        const lineLength = doc.lineAt(lineIndex).text.length;
        const character = Math.min(lineLength, change.range.start.character + change.text.length);
        const offset = doc.offsetAt(new vscode.Position(lineIndex, character));
        const syncEdit = resolveLoopIteratorSyncEdit(doc.getText(), offset);
        if (!syncEdit) { return; }

        applying = true;
        try {
            await editor.edit(
                (eb) => eb.replace(
                    new vscode.Range(
                        doc.positionAt(syncEdit.span.start),
                        doc.positionAt(syncEdit.span.end),
                    ),
                    syncEdit.newText,
                ),
                { undoStopBefore: false, undoStopAfter: false },
            );
        } finally {
            applying = false;
        }
    });

    context.subscriptions.push(sub);
}

export function registerVbaLanguageProviders(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
): VbaSymbolIndex {
    const index = new VbaSymbolIndex(bridge);

    registerVbaDiagnostics(context, index);
    registerVbaAutoBlock(context);
    registerVbaLoopIteratorSync(context);
    const docMetadata = new DocMetadataLoader();
    void docMetadata.start(context);
    registerVbaMemberCompletion(context, bridge, VBA_SELECTOR, docMetadata.registry);

    context.subscriptions.push(
        index,
        vscode.languages.registerDocumentSymbolProvider(
            VBA_SELECTOR,
            new VbaDocumentSymbolProvider(),
            { label: 'XLIDE VBA' },
        ),
        vscode.languages.registerWorkspaceSymbolProvider(
            new VbaWorkspaceSymbolProvider(index),
        ),
        vscode.languages.registerDefinitionProvider(
            VBA_SELECTOR,
            new VbaDefinitionProvider(index),
        ),
        vscode.languages.registerReferenceProvider(
            VBA_SELECTOR,
            new VbaReferenceProvider(index),
        ),
        vscode.languages.registerRenameProvider(
            VBA_SELECTOR,
            new VbaRenameProvider(index),
        ),
        vscode.languages.registerCodeActionsProvider(
            VBA_SELECTOR,
            new VbaCodeActionProvider(),
            {
                providedCodeActionKinds: [
                    vscode.CodeActionKind.QuickFix,
                    XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
                    XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
                ],
            },
        ),
        vscode.languages.registerDocumentSemanticTokensProvider(
            VBA_SELECTOR,
            new VbaTypeSemanticTokensProvider(index),
            TYPE_TOKEN_LEGEND,
        ),
        // Keep the index consistent with saves to virtual VBA documents.
        vscode.workspace.onDidSaveTextDocument((doc) => {
            if (doc.uri.scheme !== XLIDE_SCHEME) { return; }
            try {
                const { xlsmPath, moduleName } = decodeModuleUri(doc.uri);
                index.updateModuleSource(xlsmPath, moduleName, doc.getText());
            } catch {
                // Ignore URIs we cannot decode.
            }
        }),
    );

    return index;
}
