// Format Document and Format Selection for VBA modules.
//
// The formatter itself is pure (src/analyzer/format/formatModule.ts). This
// layer feeds it the document's analysis text, turns its result into one edit
// per changed line, and supplies identifier casing from the declarations the
// project index knows: a module's own symbols first, then what the project
// exports, then the VBA runtime and the host's globals and constants. VBA is
// case-insensitive, so casing never changes what code means; it only makes
// the module read the way the VBE would print it.

import * as vscode from 'vscode';
import {
    buildModuleSymbols,
    isProcedureKind,
    resolveHostConstant,
    resolveHostGlobalMember,
    resolveRuntimeConstant,
    resolveRuntimeFunction,
    resolveRuntimeObject,
    type VbaSymbol,
} from './analyzer';
import { formatVbaModule } from './analyzer/format/formatModule';
import { hostObjectModelForTokens, hostTokenForFileName } from './analyzer/host/hostRegistry';
import { hostTokensForProject } from './analyzer/host/hostLibraries';
import type { VbaProjectReference } from './vba/vbaProjectReferences';
import { analysisSourceForDocument, moduleLocationOfDocument } from './vbaDocumentLocation';
import { moduleKindFromDocument, moduleNameFromDocument } from './vbaDocumentIdentity';
import { designerHeaderEnd } from './vba/moduleSource';
import { isVbaDocument, XLIDE_SCHEME } from './xlideFileSystem';
import type { VbaProjectIndexService } from './vbaProjectIndexService';

type IdentifierCase = (name: string, offset: number) => string | undefined;

export class VbaFormattingProvider implements
    vscode.DocumentFormattingEditProvider,
    vscode.DocumentRangeFormattingEditProvider {

    constructor(
        private readonly _projects: VbaProjectIndexService,
        private readonly _log: (line: string) => void,
    ) {}

    provideDocumentFormattingEdits(
        document: vscode.TextDocument,
        options: vscode.FormattingOptions,
        token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.TextEdit[]> {
        return this._edits(document, options, undefined, token);
    }

    provideDocumentRangeFormattingEdits(
        document: vscode.TextDocument,
        range: vscode.Range,
        options: vscode.FormattingOptions,
        token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.TextEdit[]> {
        return this._edits(document, options, range, token);
    }

    private async _edits(
        document: vscode.TextDocument,
        options: vscode.FormattingOptions,
        range: vscode.Range | undefined,
        token: vscode.CancellationToken,
    ): Promise<vscode.TextEdit[]> {
        if (!isVbaDocument(document)) {
            return [];
        }
        const text = document.getText();
        // A file on disk may open with a designer block that is not VBA; the
        // analysis text blanks it and the lines it covered are left alone.
        const source = analysisSourceForDocument(document);
        const firstFormattableLine = source === text
            ? 0
            : document.positionAt(designerHeaderEnd(text)).line;
        const identifierCase = await this._identifierCase(document, source);
        if (token.isCancellationRequested) {
            return [];
        }
        const result = formatVbaModule(source, {
            tabSize: options.tabSize,
            insertSpaces: options.insertSpaces,
            identifierCase,
        });
        if (result.text === undefined) {
            this._log(`[format] declined to format ${document.uri.toString()}: ${result.refusal ?? 'unknown reason'}`);
            return [];
        }
        const before = text.split(/\r\n|\r|\n/);
        const after = result.text.split(/\r\n|\r|\n/);
        if (before.length !== after.length) {
            this._log(`[format] declined to format ${document.uri.toString()}: line count changed`);
            return [];
        }
        const edits: vscode.TextEdit[] = [];
        const fromLine = Math.max(firstFormattableLine, range?.start.line ?? 0);
        const toLine = Math.min(before.length - 1, range?.end.line ?? before.length - 1);
        for (let line = fromLine; line <= toLine; line++) {
            if (before[line] === after[line]) {
                continue;
            }
            edits.push(vscode.TextEdit.replace(
                new vscode.Range(line, 0, line, before[line].length),
                after[line],
            ));
        }
        return edits;
    }

    /**
     * Declared spellings for the module's identifiers. Locals and parameters
     * answer inside their procedure only; module-level and project-visible
     * declarations answer everywhere; the runtime and the host fill in the
     * built-ins. Every lookup is by lowercased name.
     */
    private async _identifierCase(
        document: vscode.TextDocument,
        source: string,
    ): Promise<IdentifierCase | undefined> {
        try {
            const location = moduleLocationOfDocument(document);
            const moduleName = location?.moduleName ?? moduleNameFromDocument(document);
            const moduleKind = moduleKindFromDocument(document);
            const symbols = buildModuleSymbols(moduleName, moduleKind, source);

            const procedures = (symbols.root.children ?? [])
                .filter((symbol) => isProcedureKind(symbol.kind))
                .sort((a, b) => a.fullSpan.start - b.fullSpan.start);
            const localsByProcedure = new Map<VbaSymbol, Map<string, string>>();
            const moduleLevel = new Map<string, string>();
            for (const symbol of symbols.root.children ?? []) {
                addName(moduleLevel, symbol.name);
                if (symbol.kind === 'enum') {
                    for (const member of symbol.children ?? []) {
                        addName(moduleLevel, member.name);
                    }
                }
            }

            const projectLevel = new Map<string, string>();
            let references: readonly VbaProjectReference[] = [];
            if (location && document.uri.scheme === XLIDE_SCHEME) {
                const context = await this._projects.contextForProject(location.projectPath);
                references = context.references;
                for (const symbol of context.project.visibleIdentifierSymbols(moduleName)) {
                    addName(projectLevel, symbol.name);
                }
            }

            // A project that references another application's library writes
            // that library's names too, so canonical casing has to know them.
            const hostModel = hostObjectModelForTokens(hostTokensForProject(
                location ? hostTokenForFileName(location.projectPath) : undefined,
                references,
            ));

            return (name, offset) => {
                const lower = name.toLowerCase();
                const procedure = procedureAt(procedures, offset);
                if (procedure) {
                    let locals = localsByProcedure.get(procedure);
                    if (!locals) {
                        locals = new Map();
                        for (const child of procedure.children ?? []) {
                            addName(locals, child.name);
                        }
                        localsByProcedure.set(procedure, locals);
                    }
                    const local = locals.get(lower);
                    if (local) {
                        return local;
                    }
                }
                return moduleLevel.get(lower)
                    ?? projectLevel.get(lower)
                    ?? resolveRuntimeFunction(name)?.name
                    ?? resolveRuntimeConstant(name)?.name
                    ?? resolveRuntimeObject(name)?.name
                    ?? resolveHostGlobalMember(name, hostModel)?.name
                    ?? resolveHostConstant(name, hostModel)?.name;
            };
        } catch (err) {
            // Casing is a courtesy; indentation and spacing do not wait for it.
            this._log(`[format] identifier casing unavailable: ${err instanceof Error ? err.message : String(err)}`);
            return undefined;
        }
    }
}

function addName(into: Map<string, string>, name: string): void {
    const lower = name.toLowerCase();
    if (!into.has(lower)) {
        into.set(lower, name);
    }
}

/** The procedure whose span holds `offset`, by binary search over sorted spans. */
function procedureAt(procedures: readonly VbaSymbol[], offset: number): VbaSymbol | undefined {
    let low = 0;
    let high = procedures.length - 1;
    while (low <= high) {
        const mid = (low + high) >> 1;
        const span = procedures[mid].fullSpan;
        if (offset < span.start) {
            high = mid - 1;
        } else if (offset >= span.end) {
            low = mid + 1;
        } else {
            return procedures[mid];
        }
    }
    return undefined;
}
