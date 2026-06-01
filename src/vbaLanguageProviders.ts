import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XLIDE_SCHEME, decodeModuleUri, encodeModuleUri } from './xlideFileSystem';
import { VbaSymbol, VbaSymbolIndex, VbaModuleSymbols, parseVbaModule } from './vbaSymbolIndex';
import {
    lintVbaSource,
    stripVba,
    detectSmartBlockOpener,
    findIdentifierOccurrences,
    isSmartBlockClosedAhead,
    lineStartOffsets,
    resolveLoopIteratorSyncEdit,
    smartBlockBodyText,
    VBA_IDENTIFIER_NAME_RE,
    VBA_IDENTIFIER_RE,
} from './vbaLinter';
import {
    analyzeModule,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    ModuleSymbolKind,
    normalizeDiagnosticCode,
    ProjectIndex,
    ReferenceScope,
    resolveDiagnosticCodeActions,
    resolveMemberCompletions,
    resolveTypeReferenceAt,
    resolveTypeSemanticTokens,
    SeverityOverrides,
    TypeSemanticTokenType,
    VbaDiagnostic,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    VbaSymbol as AstSymbol,
} from './analyzer';
import { registerVbaMemberCompletion } from './vbaMemberCompletion';
import { DocMetadataLoader } from './vbaDocMetadata';
import {
    buildVbaProjectIndex as buildProjectIndex,
    moduleKindFromType,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeReferenceLocations,
} from './vbaNavigation';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
];

function symbolKindToVscode(kind: VbaSymbol['kind']): vscode.SymbolKind {
    switch (kind) {
        case 'Sub': return vscode.SymbolKind.Method;
        case 'Function': return vscode.SymbolKind.Function;
        case 'PropertyGet':
        case 'PropertyLet':
        case 'PropertySet':
            return vscode.SymbolKind.Property;
        case 'Const': return vscode.SymbolKind.Constant;
        case 'Enum': return vscode.SymbolKind.Enum;
        case 'Type': return vscode.SymbolKind.Struct;
    }
}

function symbolDetail(symbol: VbaSymbol): string {
    switch (symbol.kind) {
        case 'PropertyGet': return 'Property Get';
        case 'PropertyLet': return 'Property Let';
        case 'PropertySet': return 'Property Set';
        default: return symbol.kind;
    }
}

function symbolRange(symbol: VbaSymbol): vscode.Range {
    return new vscode.Range(symbol.startLine, 0, symbol.endLine, Number.MAX_SAFE_INTEGER);
}

function selectionRange(symbol: VbaSymbol): vscode.Range {
    return new vscode.Range(
        symbol.line, symbol.column,
        symbol.line, symbol.column + symbol.length,
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
    const byModule = new Map(modules.map((m) => [m.moduleName.toLowerCase(), m]));
    byModule.set(moduleName.toLowerCase(), {
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

class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {    provideDocumentSymbols(document: vscode.TextDocument): vscode.DocumentSymbol[] {
        const symbols = parseVbaModule(document.getText());
        return symbols.map((s) => new vscode.DocumentSymbol(
            s.name,
            symbolDetail(s),
            symbolKindToVscode(s.kind),
            symbolRange(s),
            selectionRange(s),
        ));
    }
}

class VbaDefinitionProvider implements vscode.DefinitionProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideDefinition(
        document: vscode.TextDocument,
        position: vscode.Position,
    ): Promise<vscode.Location[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }

        const wordRange = document.getWordRangeAtPosition(position, VBA_IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const word = document.getText(wordRange);
        const line = document.lineAt(position.line).text;
        const qualifier = detectQualifier(line, wordRange.start.character);

        const source = document.getText();
        const { xlsmPath, moduleName } = decodeModuleUri(document.uri);
        const modules = await this._index.getAllModules(xlsmPath);
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
                const typeDefinitions = project.resolveTypeDefinitions(moduleName, typeReference.name);
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
            : project.resolveDefinition(moduleName, word, document.offsetAt(position));

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
        const modules = await this._index.getAllModules(xlsmPath);
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
            const definitions = project.resolveTypeDefinitions(moduleName, typeReference.name);
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
        const modules = await this._index.getAllModules(xlsmPath);
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
        const modules = await this._index.getAllModules(xlsmPath);
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
    return document.languageId === 'vba' || document.uri.scheme === XLIDE_SCHEME;
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

const TYPE_TOKEN_TYPES: TypeSemanticTokenType[] = [
    'class',
    'enum',
    'struct',
    'type',
];
const TYPE_TOKEN_LEGEND = new vscode.SemanticTokensLegend(TYPE_TOKEN_TYPES);

class VbaTypeSemanticTokensProvider implements vscode.DocumentSemanticTokensProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideDocumentSemanticTokens(
        document: vscode.TextDocument,
        token: vscode.CancellationToken,
    ): Promise<vscode.SemanticTokens> {
        const builder = new vscode.SemanticTokensBuilder(TYPE_TOKEN_LEGEND);
        if (!isVbaDocument(document)) { return builder.build(); }

        const source = document.getText();
        const moduleName = moduleNameFromDocument(document);
        let projectTypes: ReturnType<ProjectIndex['visibleTypeNames']> = [];
        try {
            const project = await this._projectIndexForDocument(document, source, moduleName);
            projectTypes = project.visibleTypeNames(moduleName);
        } catch {
            projectTypes = [];
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

    private async _projectIndexForDocument(
        document: vscode.TextDocument,
        source: string,
        moduleName: string,
    ): Promise<ProjectIndex> {
        if (document.uri.scheme !== XLIDE_SCHEME) {
            return buildProjectIndex([], {
                moduleName,
                moduleKind: 'standard',
                source,
            });
        }

        const decoded = decodeModuleUri(document.uri);
        const modules = await this._index.getAllModules(decoded.xlsmPath);
        const current = modules.find(
            (m) => m.moduleName.toLowerCase() === moduleName.toLowerCase(),
        );
        return buildProjectIndex(modules, {
            moduleName,
            moduleKind: moduleKindFromType(current?.type),
            source,
        });
    }
}

/**
 * Live diagnostics: structural block-balance (lintVbaSource) plus the analyzer's
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
    const timers = new Map<string, ReturnType<typeof setTimeout>>();

    const severityToVscode = (s: RuleSeverity): vscode.DiagnosticSeverity => {
        switch (s) {
            case 'error': return vscode.DiagnosticSeverity.Error;
            case 'warning': return vscode.DiagnosticSeverity.Warning;
            case 'information': return vscode.DiagnosticSeverity.Information;
            case 'hint': return vscode.DiagnosticSeverity.Hint;
        }
    };

    const run = (document: vscode.TextDocument): void => {
        void runAsync(document);
    };

    const runAsync = async (document: vscode.TextDocument): Promise<void> => {
        if (!isVbaDocument(document)) { return; }
        const config = vscode.workspace.getConfiguration('xlide');
        if (config.get<boolean>('diagnostics.enabled', true) === false) {
            collection.delete(document.uri);
            return;
        }
        const text = document.getText();
        const diagnostics: vscode.Diagnostic[] = [];

        // Project-wide names enable cross-module call and Option Explicit checks;
        // project signatures enable deterministic cross-module arity/type checks.
        // Only available for workbook-backed docs.
        const moduleName = moduleNameFromDocument(document);
        let moduleKind: ModuleSymbolKind | undefined;
        let documentType: EventHandlerDocumentType | undefined;
        let knownProcedures: ReadonlySet<string> | undefined;
        let knownIdentifiers: ReadonlySet<string> | undefined;
        let projectProcedures: ReturnType<ProjectIndex['procedureSignatures']> | undefined;
        let projectClassMembers: ReturnType<ProjectIndex['projectClassMembers']> | undefined;
        let projectTypes: ReturnType<ProjectIndex['visibleTypeNames']> | undefined;
        let knownNonTypeNames: ReturnType<ProjectIndex['visibleNonTypeNames']> | undefined;
        if (document.uri.scheme === XLIDE_SCHEME) {
            try {
                const { xlsmPath } = decodeModuleUri(document.uri);
                const modules = await index.getAllModules(xlsmPath);
                const current = modules.find(
                    (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
                );
                moduleKind = moduleKindFromType(current?.type);
                documentType = current?.documentType;
                const project = buildProjectIndex(modules, {
                    moduleName,
                    moduleKind,
                    source: text,
                });
                knownProcedures = project.visibleProcedureNames(moduleName);
                knownIdentifiers = project.visibleIdentifierNames(moduleName);
                projectProcedures = project.procedureSignatures();
                projectClassMembers = project.projectMemberSurfaces(moduleName);
                projectTypes = project.visibleTypeNames(moduleName);
                knownNonTypeNames = project.visibleNonTypeNames(moduleName);
            } catch {
                knownProcedures = undefined;
                knownIdentifiers = undefined;
                projectProcedures = undefined;
                projectClassMembers = undefined;
                projectTypes = undefined;
                knownNonTypeNames = undefined;
            }
        }

        // Structural block-balance (precise per-line spans).
        for (const p of lintVbaSource(text)) {
            const diag = new vscode.Diagnostic(
                new vscode.Range(p.line, p.startCol, p.line, p.endCol),
                p.message,
                p.severity === 'error'
                    ? vscode.DiagnosticSeverity.Error
                    : vscode.DiagnosticSeverity.Warning,
            );
            diag.source = 'XLIDE';
            if (p.code) {
                diag.code = p.code;
            }
            diagnostics.push(diag);
        }

        // Semantic rules (offset spans -> ranges).
        const optionExplicit = config.get<string>('diagnostics.optionExplicit', 'warning');
        const severities: SeverityOverrides = {
            optionExplicitMissing:
                optionExplicit === 'off' ? 'off' : (optionExplicit as RuleSeverity),
        };
        let semantic: VbaDiagnostic[] = [];
        try {
            semantic = analyzeModule(text, {
                moduleName,
                moduleKind,
                documentType,
                severities,
                knownProcedures,
                knownIdentifiers,
                projectProcedures,
                projectClassMembers,
                projectTypes,
                knownNonTypeNames,
            });
        } catch {
            semantic = [];
        }
        for (const d of semantic) {
            const diag = new vscode.Diagnostic(
                new vscode.Range(
                    document.positionAt(d.span.start),
                    document.positionAt(d.span.end),
                ),
                d.message,
                severityToVscode(d.severity),
            );
            diag.source = 'XLIDE';
            diag.code = d.code;
            diagnostics.push(diag);
        }

        collection.set(document.uri, diagnostics);
    };

    const schedule = (document: vscode.TextDocument): void => {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const existing = timers.get(key);
        if (existing) { clearTimeout(existing); }
        timers.set(key, setTimeout(() => {
            timers.delete(key);
            run(document);
        }, 300));
    };

    context.subscriptions.push(
        collection,
        vscode.workspace.onDidOpenTextDocument(run),
        vscode.workspace.onDidChangeTextDocument((e) => schedule(e.document)),
        vscode.workspace.onDidCloseTextDocument((doc) => {
            const key = doc.uri.toString();
            const t = timers.get(key);
            if (t) { clearTimeout(t); timers.delete(key); }
            collection.delete(doc.uri);
        }),
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('xlide.diagnostics')) {
                vscode.workspace.textDocuments.forEach(run);
            }
        }),
    );
    vscode.workspace.textDocuments.forEach(run);
}

class VbaCodeActionProvider implements vscode.CodeActionProvider {
    public provideCodeActions(
        document: vscode.TextDocument,
        _range: vscode.Range,
        context: vscode.CodeActionContext,
        _token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.CodeAction[]> {
        if (!isVbaDocument(document)) { return []; }
        if (
            context.only &&
            !context.only.contains(vscode.CodeActionKind.QuickFix) &&
            !vscode.CodeActionKind.QuickFix.contains(context.only)
        ) {
            return [];
        }

        const source = document.getText();
        const actions: vscode.CodeAction[] = [];
        let lintProblems: ReturnType<typeof lintVbaSource> | undefined;
        for (const diagnostic of context.diagnostics) {
            if (diagnostic.source !== 'XLIDE') { continue; }
            const code = normalizeDiagnosticCode(diagnostic.code);
            if (!code) { continue; }
            const lintProblem = code === 'missing-block-closer'
                ? matchingLintProblem(
                    diagnostic,
                    lintProblems ??= lintVbaSource(source),
                )
                : undefined;
            const fixes = resolveDiagnosticCodeActions(source, {
                code,
                message: diagnostic.message,
                expectedClose: lintProblem?.expectedClose,
                insertLine: lintProblem?.insertLine,
                span: {
                    start: document.offsetAt(diagnostic.range.start),
                    end: document.offsetAt(diagnostic.range.end),
                },
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

function matchingLintProblem(
    diagnostic: vscode.Diagnostic,
    problems: ReturnType<typeof lintVbaSource>,
): ReturnType<typeof lintVbaSource>[number] | undefined {
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
        const opener = detectSmartBlockOpener(stripVba(openerLine));
        if (!opener) { return; }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = doc.getText().split(/\r\n|\r|\n/).map(stripVba);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const indent = /^[ \t]*/.exec(openerLine)?.[0] ?? '';
        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const bodyLine = doc.lineAt(bodyLineIndex).text;
        if (!/^[ \t]*$/.test(bodyLine)) { return; }
        const bodyText = smartBlockBodyText(openerLine, bodyLine, opener);
        const bodyRange = new vscode.Range(
            new vscode.Position(bodyLineIndex, 0),
            new vscode.Position(bodyLineIndex, bodyLine.length),
        );

        applying = true;
        try {
            await editor.edit(
                (eb) => eb.replace(
                    bodyRange,
                    closedAhead
                        ? bodyText
                        : `${bodyText}${eol}${indent}${opener.endKeyword}`,
                ),
                { undoStopBefore: false, undoStopAfter: true },
            );
        } finally {
            applying = false;
        }

        // Keep the caret on the indented body line, above the inserted End.
        const caret = new vscode.Position(
            bodyLineIndex,
            bodyText.length,
        );
        editor.selection = new vscode.Selection(caret, caret);
    });

    context.subscriptions.push(sub);
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
            { providedCodeActionKinds: [vscode.CodeActionKind.QuickFix] },
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
                void index.refreshModule(xlsmPath, moduleName);
            } catch {
                // Ignore URIs we cannot decode.
            }
        }),
    );

    return index;
}
