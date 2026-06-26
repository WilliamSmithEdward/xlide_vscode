// Pure, vscode-free reference/rename resolution shared by the VBA reference and
// rename providers. Returns offset-based reference spans; the provider layer
// converts them to vscode.Location / WorkspaceEdit.
//
// A reference to a symbol is the union of:
//   - bare-identifier occurrences that bind to it (scope- and shadow-aware,
//     narrowed to the home module for object-module members whose bare calls
//     only bind in-module), and
//   - member-access / module-qualified occurrences (`Module.Proc`, `obj.Member`,
//     `Me.Member`) that semantically resolve to its declaration.
//
// Member definitions are derived from BOTH the invocation site and the
// bare-resolution result, so the same complete set is returned whether the user
// invoked from the declaration, a qualified reference, or a bare call site.

import {
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    precededByMemberAccessDot,
    ProjectIndex,
    ReferenceScope,
    resolveMemberDefinitionsAt,
    tokenizeCached,
    type MemberCompletionContext,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    type VbaSymbol as AstSymbol,
} from './analyzer';
import { moduleKindFromType } from './vbaProjectAnalysis';
import { findIdentifierOccurrences, lineStartOffsets } from './vbaSourceScan';
import type { VbaModuleSymbols } from './vbaSymbolIndex';

/** An offset-based reference location within one workbook module. */
export interface ReferenceSpan {
    /** Original-cased owning module name. */
    moduleName: string;
    /** Zero-based start line. */
    line: number;
    /** Zero-based start column. */
    column: number;
    /** Length of the matched identifier. */
    length: number;
}

export interface SymbolReferenceResult {
    references: ReferenceSpan[];
    /** True when a renameable/findable member or scope symbol resolved. */
    hasSymbol: boolean;
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

/**
 * Source-backed member definitions at `memberEndOffset` (a `recv.Member`
 * reference or a member declaration), or empty when the cursor is not on a
 * member.
 */
export function sourceMemberDefinitionsAt(
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

/** The member whose declaration nameSpan covers `offset`, if any. */
export function projectClassMemberAtDefinition(
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

function offsetToLineColumn(source: string, offset: number): { line: number; column: number } {
    const starts = lineStartOffsets(source);
    let line = 0;
    for (let i = 0; i < starts.length; i += 1) {
        if (starts[i] <= offset) { line = i; } else { break; }
    }
    return { line, column: offset - (starts[line] ?? 0) };
}

function memberDefinitionKey(definition: VbaProjectClassMemberDefinition): string {
    return `${definition.moduleName.toLowerCase()}:${definition.nameSpan.start}`;
}

function mergeMemberDefinitions(
    a: readonly VbaProjectClassMemberDefinition[],
    b: readonly VbaProjectClassMemberDefinition[],
): VbaProjectClassMemberDefinition[] {
    const out: VbaProjectClassMemberDefinition[] = [];
    const seen = new Set<string>();
    for (const def of [...a, ...b]) {
        const key = memberDefinitionKey(def);
        if (!seen.has(key)) {
            seen.add(key);
            out.push(def);
        }
    }
    return out;
}

/**
 * Member-surface definitions for a bare-resolved symbol. Lets a rename/find
 * invoked at a bare call site still locate the symbol's module-qualified or
 * member-access references, by matching each resolved declaration span against
 * the project member surfaces.
 */
function projectMemberDefinitionsForSymbols(
    project: ProjectIndex,
    moduleName: string,
    name: string,
    scopeDefinitions: readonly AstSymbol[],
): VbaProjectClassMemberDefinition[] {
    if (scopeDefinitions.length === 0) { return []; }
    const wanted = new Set(
        scopeDefinitions.map((d) => `${d.moduleName.toLowerCase()}:${d.nameSpan.start}`),
    );
    const lower = name.toLowerCase();
    const out: VbaProjectClassMemberDefinition[] = [];
    const seen = new Set<string>();
    for (const surface of project.projectMemberSurfaces(moduleName)) {
        for (const member of surface.members) {
            if (member.name.toLowerCase() !== lower) { continue; }
            for (const def of member.definitions ?? []) {
                const key = memberDefinitionKey(def);
                if (wanted.has(key) && !seen.has(key)) {
                    seen.add(key);
                    out.push(def);
                }
            }
        }
    }
    return out;
}

/**
 * Member-access / module-qualified references that semantically resolve to one
 * of `definitions`, plus (when `includeDeclaration`) the declarations
 * themselves. Bare occurrences fall through to {@link bareReferences}.
 */
function memberAccessReferences(
    byModule: Map<string, VbaModuleSymbols>,
    project: ProjectIndex,
    modules: VbaModuleSymbols[],
    memberName: string,
    definitions: readonly VbaProjectClassMemberDefinition[],
    includeDeclaration: boolean,
): ReferenceSpan[] {
    const targetKeys = new Set(definitions.map(memberDefinitionKey));
    const seen = new Set<string>();
    const out: ReferenceSpan[] = [];
    const push = (span: ReferenceSpan): void => {
        const key = `${span.moduleName.toLowerCase()}:${span.line}:${span.column}`;
        if (!seen.has(key)) {
            seen.add(key);
            out.push(span);
        }
    };

    if (includeDeclaration) {
        for (const definition of definitions) {
            const mod = byModule.get(definition.moduleName.toLowerCase());
            if (!mod) { continue; }
            const pos = offsetToLineColumn(mod.source, definition.nameSpan.start);
            push({
                moduleName: mod.moduleName,
                line: pos.line,
                column: pos.column,
                length: definition.nameSpan.end - definition.nameSpan.start,
            });
        }
    }

    const codeNames = codeNamesForModules(modules);
    for (const mod of byModule.values()) {
        const occurrences = findIdentifierOccurrences(mod.source, memberName);
        if (occurrences.length === 0) { continue; }
        const ctx: MemberCompletionContext = {
            codeNames,
            meType: meHostTypeForModule(mod.moduleName, mod.type, mod.documentType),
            meProjectType: meProjectTypeForModule(mod.moduleName, mod.type),
            projectClassMembers: project.projectMemberSurfaces(mod.moduleName),
        };
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
            push({ moduleName: mod.moduleName, line: occ.line, column: occ.column, length: memberName.length });
        }
    }
    return out;
}

function symbolKey(symbol: AstSymbol): string {
    return `${symbol.moduleName.toLowerCase()}:${symbol.nameSpan.start}`;
}

/**
 * Bare-identifier references that bind to one of `targetDefinitions`. Each
 * textual occurrence is resolved in its own module context and kept only when it
 * binds *solely* to the target declarations, which makes the bare pass as
 * precise as the member pass: local/parameter shadows resolve to the shadowing
 * symbol and are dropped, a same-named member or procedure in another module
 * resolves to that other symbol and is dropped, and an ambiguous bare call
 * (resolving to the target *and* an unrelated same-named export) is left alone
 * rather than guessed. Occurrences preceded by a member-access dot (`recv.word`)
 * are owned by {@link memberAccessReferences}.
 */
function bareReferences(
    byModule: Map<string, VbaModuleSymbols>,
    project: ProjectIndex,
    scope: ReferenceScope,
    word: string,
    targetDefinitions: readonly AstSymbol[],
): ReferenceSpan[] {
    const targetKeys = new Set(targetDefinitions.map(symbolKey));
    if (targetKeys.size === 0) { return []; }
    const out: ReferenceSpan[] = [];
    for (const moduleName of scope.searchModules) {
        const mod = byModule.get(moduleName.toLowerCase());
        if (!mod) { continue; }
        for (const occ of findIdentifierOccurrences(mod.source, word)) {
            if (precededByMemberAccessDot(mod.source, occ.offset)) {
                continue;
            }
            const resolved = project.resolveBareIdentifier(mod.moduleName, word, occ.offset, 'expression');
            if (
                resolved.definitions.length === 0 ||
                !resolved.definitions.every((d) => targetKeys.has(symbolKey(d)))
            ) {
                continue;
            }
            out.push({ moduleName: mod.moduleName, line: occ.line, column: occ.column, length: word.length });
        }
    }
    return out;
}

function dedupeReferences(spans: readonly ReferenceSpan[]): ReferenceSpan[] {
    const seen = new Set<string>();
    const out: ReferenceSpan[] = [];
    for (const span of spans) {
        const key = `${span.moduleName.toLowerCase()}:${span.line}:${span.column}`;
        if (!seen.has(key)) {
            seen.add(key);
            out.push(span);
        }
    }
    return out;
}

/**
 * Unified reference/rename resolution. See the file header for the model.
 */
export function collectSymbolReferences(
    byModule: Map<string, VbaModuleSymbols>,
    project: ProjectIndex,
    modules: VbaModuleSymbols[],
    source: string,
    moduleName: string,
    current: VbaModuleSymbols | undefined,
    word: string,
    wordEndOffset: number,
    positionOffset: number,
    includeDeclaration: boolean,
): SymbolReferenceResult {
    const entryMemberDefinitions = sourceMemberDefinitionsAt(
        source,
        word,
        wordEndOffset,
        project,
        modules,
        moduleName,
        current?.type,
        current?.documentType,
    );
    const memberAtDefinition = entryMemberDefinitions.length > 0
        ? undefined
        : projectClassMemberAtDefinition(project, moduleName, word, positionOffset);
    const invocationMemberDefinitions = entryMemberDefinitions.length > 0
        ? entryMemberDefinitions
        : memberAtDefinition?.definitions ?? [];

    const scope = project.referenceScope(moduleName, word, positionOffset);

    const targetMemberDefinitions = mergeMemberDefinitions(
        invocationMemberDefinitions,
        projectMemberDefinitionsForSymbols(project, moduleName, word, scope.definitions),
    );

    if (targetMemberDefinitions.length === 0 && scope.definitions.length === 0) {
        return { references: [], hasSymbol: false };
    }

    const memberSpans = targetMemberDefinitions.length > 0
        ? memberAccessReferences(byModule, project, modules, word, targetMemberDefinitions, includeDeclaration)
        : [];

    const bareSpans = bareReferences(byModule, project, scope, word, scope.definitions);

    let references = dedupeReferences([...memberSpans, ...bareSpans]);

    if (!includeDeclaration) {
        const declKeys = new Set<string>([
            ...scope.definitions.map((d) => `${d.moduleName.toLowerCase()}:${d.nameSpan.start}`),
            ...targetMemberDefinitions.map(memberDefinitionKey),
        ]);
        references = references.filter((span) => {
            const mod = byModule.get(span.moduleName.toLowerCase());
            if (!mod) { return true; }
            const offset = (lineStartOffsets(mod.source)[span.line] ?? 0) + span.column;
            return !declKeys.has(`${span.moduleName.toLowerCase()}:${offset}`);
        });
    }

    return { references, hasSymbol: true };
}
