import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import { XLIDE_SCHEME, decodeModuleUri, encodeModuleUri } from './xlideFileSystem';
import { VbaSymbol, VbaSymbolIndex, parseVbaModule } from './vbaSymbolIndex';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
];

const IDENTIFIER_RE = /[A-Za-z_][A-Za-z0-9_]*/;
const WORD_RE_GLOBAL = /[A-Za-z_][A-Za-z0-9_]*/g;

function symbolKindToVscode(kind: VbaSymbol['kind']): vscode.SymbolKind {
    switch (kind) {
        case 'Sub': return vscode.SymbolKind.Method;
        case 'Function': return vscode.SymbolKind.Function;
        case 'PropertyGet':
        case 'PropertyLet':
        case 'PropertySet':
            return vscode.SymbolKind.Property;
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

/**
 * Finds all whole-word occurrences of `name` in `source`, ignoring
 * matches inside string literals and line comments.
 */
function findIdentifierOccurrences(
    source: string,
    name: string,
): Array<{ line: number; column: number }> {
    const lines = source.split(/\r\n|\r|\n/);
    const lower = name.toLowerCase();
    const out: Array<{ line: number; column: number }> = [];
    for (let i = 0; i < lines.length; i++) {
        const raw = lines[i];
        const stripped = stripStringsAndComment(raw);
        WORD_RE_GLOBAL.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = WORD_RE_GLOBAL.exec(stripped)) !== null) {
            if (m[0].toLowerCase() === lower) {
                out.push({ line: i, column: m.index });
            }
        }
    }
    return out;
}

/** Replace string contents and trailing comment with spaces so columns stay aligned. */
function stripStringsAndComment(line: string): string {
    const chars = line.split('');
    let inString = false;
    for (let i = 0; i < chars.length; i++) {
        const c = chars[i];
        if (inString) {
            if (c === '"') {
                if (chars[i + 1] === '"') {
                    chars[i] = ' ';
                    chars[i + 1] = ' ';
                    i++;
                    continue;
                }
                inString = false;
                continue;
            }
            chars[i] = ' ';
            continue;
        }
        if (c === '"') {
            inString = true;
            continue;
        }
        if (c === "'") {
            for (let j = i; j < chars.length; j++) { chars[j] = ' '; }
            break;
        }
    }
    return chars.join('');
}

// ---------------------------------------------------------------------------
// Provider implementations
// ---------------------------------------------------------------------------

class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    provideDocumentSymbols(document: vscode.TextDocument): vscode.DocumentSymbol[] {
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

        const wordRange = document.getWordRangeAtPosition(position, IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const word = document.getText(wordRange);
        const line = document.lineAt(position.line).text;
        const qualifier = detectQualifier(line, wordRange.start.character);

        const { xlsmPath, moduleName: currentModule } = decodeModuleUri(document.uri);
        const matches: vscode.Location[] = [];

        const modules = await this._index.getAllModules(xlsmPath);
        for (const mod of modules) {
            if (qualifier && unquoteModule(qualifier).toLowerCase() !== mod.moduleName.toLowerCase()) {
                continue;
            }
            for (const sym of mod.symbols) {
                if (sym.name.toLowerCase() !== word.toLowerCase()) { continue; }
                if (!qualifier && !sym.isPublic && mod.moduleName !== currentModule) {
                    continue;
                }
                matches.push(new vscode.Location(
                    encodeModuleUri(xlsmPath, mod.moduleName),
                    selectionRange(sym),
                ));
            }
        }
        return matches.length > 0 ? matches : undefined;
    }
}

class VbaReferenceProvider implements vscode.ReferenceProvider {
    constructor(private readonly _index: VbaSymbolIndex) {}

    async provideReferences(
        document: vscode.TextDocument,
        position: vscode.Position,
    ): Promise<vscode.Location[] | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        const wordRange = document.getWordRangeAtPosition(position, IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const word = document.getText(wordRange);

        const { xlsmPath } = decodeModuleUri(document.uri);
        const modules = await this._index.getAllModules(xlsmPath);
        const locations: vscode.Location[] = [];
        for (const mod of modules) {
            const uri = encodeModuleUri(xlsmPath, mod.moduleName);
            for (const occ of findIdentifierOccurrences(mod.source, word)) {
                locations.push(new vscode.Location(
                    uri,
                    new vscode.Range(occ.line, occ.column, occ.line, occ.column + word.length),
                ));
            }
        }
        return locations;
    }
}

class VbaRenameProvider implements vscode.RenameProvider {
    constructor(
        private readonly _index: VbaSymbolIndex,
        private readonly _bridge: PythonBridge,
    ) {}

    async prepareRename(
        document: vscode.TextDocument,
        position: vscode.Position,
    ): Promise<vscode.Range | { range: vscode.Range; placeholder: string }> {
        if (document.uri.scheme !== XLIDE_SCHEME) {
            throw new Error('Rename is only supported in XLIDE VBA modules.');
        }
        const wordRange = document.getWordRangeAtPosition(position, IDENTIFIER_RE);
        if (!wordRange) { throw new Error('No symbol at cursor.'); }
        const word = document.getText(wordRange);

        const { xlsmPath } = decodeModuleUri(document.uri);
        const modules = await this._index.getAllModules(xlsmPath);
        const found = modules.some((m) =>
            m.symbols.some((s) => s.name.toLowerCase() === word.toLowerCase()),
        );
        if (!found) {
            throw new Error(`'${word}' is not a known VBA procedure in this workbook.`);
        }
        return { range: wordRange, placeholder: word };
    }

    async provideRenameEdits(
        document: vscode.TextDocument,
        position: vscode.Position,
        newName: string,
    ): Promise<vscode.WorkspaceEdit | undefined> {
        if (document.uri.scheme !== XLIDE_SCHEME) { return undefined; }
        if (!IDENTIFIER_RE.test(newName) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(newName)) {
            throw new Error(`'${newName}' is not a valid VBA identifier.`);
        }
        const wordRange = document.getWordRangeAtPosition(position, IDENTIFIER_RE);
        if (!wordRange) { return undefined; }
        const oldName = document.getText(wordRange);
        if (oldName.toLowerCase() === newName.toLowerCase()) { return undefined; }

        const { xlsmPath } = decodeModuleUri(document.uri);
        const modules = await this._index.getAllModules(xlsmPath);

        const edit = new vscode.WorkspaceEdit();
        for (const mod of modules) {
            const occs = findIdentifierOccurrences(mod.source, oldName);
            if (occs.length === 0) { continue; }
            const uri = encodeModuleUri(xlsmPath, mod.moduleName);
            for (const occ of occs) {
                edit.replace(
                    uri,
                    new vscode.Range(occ.line, occ.column, occ.line, occ.column + oldName.length),
                    newName,
                );
            }
        }
        return edit;
    }
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

export function registerVbaLanguageProviders(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
): VbaSymbolIndex {
    const index = new VbaSymbolIndex(bridge);

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
            new VbaRenameProvider(index, bridge),
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
