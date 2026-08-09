// Pure, dependency-free structural block-balance diagnostics used by the
// diagnostics provider. Keeping this module free of any `vscode` import means
// it can be unit-tested directly with vitest.

import { leadingWhitespace, stripVba, toLogicalLines } from './vbaSourceScan';

export interface VbaStructuralDiagnostic {
    /** 0-based physical line of the relevant token. */
    line: number;
    /** 0-based start column (inclusive). */
    startCol: number;
    /** 0-based end column (exclusive). */
    endCol: number;
    /** Stable diagnostic code for editor integrations. */
    code?: 'missing-block-closer' | 'unmatched-block-closer' | 'mismatched-end-keyword' | 'module-declaration-in-procedure';
    /** Closing phrase that can repair a missing-block diagnostic. */
    expectedClose?: string;
    /** 0-based physical line before which the missing closer should be inserted. */
    insertLine?: number;
    /** Token replacement that can repair an obvious mismatched closer. */
    expectedCloseReplacement?: {
        line: number;
        startCol: number;
        endCol: number;
        text: string;
    };
    message: string;
    severity: 'error' | 'warning';
}

export interface VbaStructuralAnalysisOptions {
    /**
     * Returns true for physical source lines that belong to a proven-inactive
     * conditional-compilation branch. Preprocessor lines are still analyzed so
     * #If/#Else/#End If balance remains checked.
     */
    isInactiveLine?: (line: number) => boolean;
}

export type BlockKind =
    | 'Sub' | 'Function' | 'Property'
    | 'If' | 'With' | 'Select' | 'Type' | 'Enum'
    | 'For' | 'Do' | 'While' | 'PreprocessorIf';

export interface OpenBlock {
    kind: BlockKind;
    /** 0-based physical line of the opener. */
    line: number;
    /** Friendly descriptor, e.g. "Sub Foo" or "If". */
    label: string;
}

interface ColumnSpan {
    startCol: number;
    endCol: number;
}

interface ModuleDeclarationInProcedureHit {
    label: string;
    span: ColumnSpan;
}

/** The closing phrase expected for each block kind. */
export const CLOSE_PHRASE: Record<BlockKind, string> = {
    Sub: 'End Sub',
    Function: 'End Function',
    Property: 'End Property',
    If: 'End If',
    With: 'End With',
    Select: 'End Select',
    Type: 'End Type',
    Enum: 'End Enum',
    For: 'Next',
    Do: 'Loop',
    While: 'Wend',
    PreprocessorIf: '#End If',
};

/** The opener keyword shown when a stray closer has no match. */
const OPEN_WORD: Record<BlockKind, string> = {
    Sub: 'Sub',
    Function: 'Function',
    Property: 'Property',
    If: 'If',
    With: 'With',
    Select: 'Select Case',
    Type: 'Type',
    Enum: 'Enum',
    For: 'For',
    Do: 'Do',
    While: 'While',
    PreprocessorIf: '#If',
};

/** Detects a block closer on a stripped, trimmed logical line. */
export function matchCloser(t: string): BlockKind | undefined {
    if (/^#\s*End\s*If\b/i.test(t) || /^#\s*EndIf\b/i.test(t)) { return 'PreprocessorIf'; }
    const end = /^End\s+(Sub|Function|Property|If|With|Select|Type|Enum)\b/i.exec(t);
    if (end) {
        const w = end[1].toLowerCase();
        if (w === 'sub') { return 'Sub'; }
        if (w === 'function') { return 'Function'; }
        if (w === 'property') { return 'Property'; }
        if (w === 'if') { return 'If'; }
        if (w === 'with') { return 'With'; }
        if (w === 'select') { return 'Select'; }
        if (w === 'type') { return 'Type'; }
        return 'Enum';
    }
    if (/^Loop\b/i.test(t)) { return 'Do'; }
    if (/^Wend\b/i.test(t)) { return 'While'; }
    return undefined;
}

/** Detects a block opener on a stripped, trimmed logical line. */
export function matchOpener(t: string): OpenBlock | undefined {
    let m: RegExpExecArray | null;
    if (/^#\s*If\b/i.test(t)) { return { kind: 'PreprocessorIf', line: 0, label: '#If' }; }
    // \p{L}: VBA procedure/type names may use any locale letter; the old
    // ASCII class made `Sub Proverka()` written in Cyrillic invisible here,
    // so its End Sub reported as unmatched - a false error on Russian code.
    // \p{M} because Thai and Devanagari names are a base letter plus a
    // combining mark, and truncating at the mark captured only half the name.
    //
    // The class deliberately does NOT require a valid identifier START here.
    // Block matching asks whether a block opened, and `Public Sub 1Bad()`
    // opens one: the name is wrong, which invalid-identifier-start already
    // reports on that very line. Refusing to see the header made its `End
    // Sub` report as an orphan too - two findings for one mistake, the second
    // naming a line where nothing is wrong.
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\s+([\p{L}\p{M}\p{N}_]+)/iu.exec(t);
    if (m) {
        const kind = (/^sub$/i.test(m[1]) ? 'Sub' : 'Function') as BlockKind;
        return { kind, line: 0, label: `${kind} ${m[2]}` };
    }
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Property\s+(Get|Let|Set)\s+([\p{L}\p{M}\p{N}_]+)/iu.exec(t);
    if (m) { return { kind: 'Property', line: 0, label: `Property ${m[2]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Type\s+([\p{L}\p{M}\p{N}_]+)/iu.exec(t);
    if (m) { return { kind: 'Type', line: 0, label: `Type ${m[1]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Enum\s+([\p{L}\p{M}\p{N}_]+)/iu.exec(t);
    if (m) { return { kind: 'Enum', line: 0, label: `Enum ${m[1]}` }; }
    if (/^Select\s+Case\b/i.test(t)) { return { kind: 'Select', line: 0, label: 'Select Case' }; }
    if (/^If\b/i.test(t) && /\bThen\s*$/i.test(t)) { return { kind: 'If', line: 0, label: 'If' }; }
    if (/^For\b/i.test(t)) { return { kind: 'For', line: 0, label: 'For' }; }
    if (/^Do\b/i.test(t)) { return { kind: 'Do', line: 0, label: 'Do' }; }
    if (/^While\b/i.test(t)) { return { kind: 'While', line: 0, label: 'While' }; }
    if (/^With\b/i.test(t)) { return { kind: 'With', line: 0, label: 'With' }; }
    return undefined;
}

function fullLineProblem(
    physical: string[], line: number, message: string, severity: 'error' | 'warning',
    details: Pick<VbaStructuralDiagnostic, 'code' | 'expectedClose' | 'insertLine' | 'expectedCloseReplacement'> = {},
    span = defaultDiagnosticColumnSpan(physical[line] ?? ''),
): VbaStructuralDiagnostic {
    return {
        line,
        startCol: span.startCol,
        endCol: span.endCol,
        message,
        severity,
        ...details,
    };
}

function defaultDiagnosticColumnSpan(raw: string): ColumnSpan {
    const startCol = raw.length - raw.trimStart().length;
    return {
        startCol,
        endCol: Math.max(raw.length, startCol + 1),
    };
}

function blockOpenerColumnSpan(raw: string, kind: BlockKind): ColumnSpan | undefined {
    const stripped = stripVba(raw);
    switch (kind) {
        case 'Sub':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub)\b/i);
        case 'Function':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Function)\b/i);
        case 'Property':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Property\s+(?:Get|Let|Set))\b/i);
        case 'If':
            return capturedColumnSpan(stripped, /^(\s*)(If)\b/i);
        case 'With':
            return capturedColumnSpan(stripped, /^(\s*)(With)\b/i);
        case 'Select':
            return capturedColumnSpan(stripped, /^(\s*)(Select\s+Case)\b/i);
        case 'Type':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Global)\s+)?(Type)\b/i);
        case 'Enum':
            return capturedColumnSpan(stripped, /^(\s*)(?:(?:Public|Private|Global)\s+)?(Enum)\b/i);
        case 'For':
            return capturedColumnSpan(stripped, /^(\s*)(For(?:\s+Each)?)\b/i);
        case 'Do':
            return capturedColumnSpan(stripped, /^(\s*)(Do(?:\s+(?:While|Until))?)\b/i);
        case 'While':
            return capturedColumnSpan(stripped, /^(\s*)(While)\b/i);
        case 'PreprocessorIf':
            return capturedColumnSpan(stripped, /^(\s*)(#\s*If)\b/i);
    }
}

function blockCloserColumnSpan(raw: string, kind: BlockKind): ColumnSpan | undefined {
    const stripped = stripVba(raw);
    switch (kind) {
        case 'Sub':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Sub)\b/i);
        case 'Function':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Function)\b/i);
        case 'Property':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Property)\b/i);
        case 'If':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+If)\b/i);
        case 'With':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+With)\b/i);
        case 'Select':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Select)\b/i);
        case 'Type':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Type)\b/i);
        case 'Enum':
            return capturedColumnSpan(stripped, /^(\s*)(End\s+Enum)\b/i);
        case 'For':
            return capturedColumnSpan(stripped, /^(\s*)(Next)\b/i);
        case 'Do':
            return capturedColumnSpan(stripped, /^(\s*)(Loop)\b/i);
        case 'While':
            return capturedColumnSpan(stripped, /^(\s*)(Wend)\b/i);
        case 'PreprocessorIf':
            return capturedColumnSpan(stripped, /^(\s*)(#\s*End\s*If|#\s*EndIf)\b/i);
    }
}

function preprocessorBranchColumnSpan(raw: string): ColumnSpan | undefined {
    return capturedColumnSpan(stripVba(raw), /^(\s*)(#\s*ElseIf|#\s*Else)\b/i);
}

function moduleDeclarationInProcedureHit(raw: string): ModuleDeclarationInProcedureHit | undefined {
    const stripped = stripVba(raw);
    const procedure = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\b/i,
    );
    if (procedure) {
        return { label: 'Procedure declarations', span: procedure };
    }

    const property = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Property\s+(?:Get|Let|Set))\b/i,
    );
    if (property) {
        return { label: 'Procedure declarations', span: property };
    }

    const typeDeclaration = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Global)\s+)?(Type)\b/i,
    );
    if (typeDeclaration) {
        return { label: 'Type declarations', span: typeDeclaration };
    }

    const enumDeclaration = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private|Global)\s+)?(Enum)\b/i,
    );
    if (enumDeclaration) {
        return { label: 'Enum declarations', span: enumDeclaration };
    }

    const declare = capturedColumnSpan(
        stripped,
        /^(\s*)(?:(?:Public|Private)\s+)?(Declare)\b/i,
    );
    if (declare) {
        return { label: 'Declare statements', span: declare };
    }

    return undefined;
}

function activeProcedureBlock(stack: OpenBlock[]): OpenBlock | undefined {
    for (let i = stack.length - 1; i >= 0; i--) {
        const block = stack[i];
        if (block.kind === 'Sub' || block.kind === 'Function' || block.kind === 'Property') {
            return block;
        }
    }
    return undefined;
}

function isProcedureBlockKind(kind: BlockKind): boolean {
    return kind === 'Sub' || kind === 'Function' || kind === 'Property';
}

function isConditionalAlternativeProcedureHeader(
    stack: readonly OpenBlock[],
    opener: OpenBlock,
    preprocessorDepth: number,
): boolean {
    if (preprocessorDepth === 0 || !isProcedureBlockKind(opener.kind)) {
        return false;
    }
    const active = stack[stack.length - 1];
    return !!active &&
        active.kind === opener.kind &&
        active.label.toLowerCase() === opener.label.toLowerCase();
}

function isTypeFieldNamedTypeInsideType(stack: readonly OpenBlock[], opener: OpenBlock, text: string): boolean {
    return opener.kind === 'Type' &&
        stack[stack.length - 1]?.kind === 'Type' &&
        /^Type\s+As\b/i.test(text.trim());
}

function isPreprocessorLine(trimmed: string): boolean {
    return /^#\s*(?:Const|If|ElseIf|Else|End\s*If|EndIf)\b/i.test(trimmed);
}

function capturedColumnSpan(stripped: string, pattern: RegExp): ColumnSpan | undefined {
    const match = pattern.exec(stripped);
    if (!match) {
        return undefined;
    }
    const prefix = match[1] ?? '';
    const text = match[2] ?? match[0].slice(prefix.length);
    const startCol = prefix.length + match[0].slice(prefix.length).indexOf(text);
    return {
        startCol,
        endCol: startCol + text.length,
    };
}

/**
 * Performs structural block-balance analysis on VBA source, reporting:
 *  - openers with no matching closer (missing `End Sub`, `Next`, ...),
 *  - closers with no matching opener (stray `End If`, `Loop`, ...),
 *  - mismatched nesting (an inner block left unclosed).
 */
export function analyzeVbaStructure(
    source: string,
    options: VbaStructuralAnalysisOptions = {},
): VbaStructuralDiagnostic[] {
    const physical = source.split(/\r\n|\r|\n/);
    const { logical } = toLogicalLines(source);
    const stack: OpenBlock[] = [];
    const preprocessorStack: number[] = [];
    const problems: VbaStructuralDiagnostic[] = [];

    const closeOne = (closerKind: BlockKind, line: number, closerWord: string): void => {
        let idx = -1;
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].kind === closerKind) { idx = k; break; }
        }
        if (idx === -1) {
            const top = stack[stack.length - 1];
            if (top && isProcedureBlockKind(top.kind) && isProcedureBlockKind(closerKind)) {
                // VBE accepts End Sub/Function/Property interchangeably as
                // procedure closers (oracle-verified: the mismatch compiles), so
                // this is a style warning, not a missing-closer compile error. The
                // procedure is still closed (stack.pop below); the quick-fix swaps
                // the keyword to match the opener.
                const closerSpan = blockCloserColumnSpan(physical[line] ?? '', closerKind);
                problems.push(fullLineProblem(
                    physical, top.line,
                    `Procedure '${top.label}' is closed with '${CLOSE_PHRASE[closerKind]}'; use '${CLOSE_PHRASE[top.kind]}' to match the opening keyword.`,
                    'warning',
                    {
                        code: 'mismatched-end-keyword',
                        expectedClose: CLOSE_PHRASE[top.kind],
                        insertLine: line,
                        expectedCloseReplacement: closerSpan
                            ? {
                                line,
                                startCol: closerSpan.startCol,
                                endCol: closerSpan.endCol,
                                text: CLOSE_PHRASE[top.kind],
                            }
                            : undefined,
                    },
                    blockOpenerColumnSpan(physical[top.line] ?? '', top.kind),
                ));
                stack.pop();
                return;
            }
            problems.push(fullLineProblem(
                physical, line,
                `'${closerWord}' has no matching '${OPEN_WORD[closerKind]}'.`,
                'error',
                { code: 'unmatched-block-closer' },
                blockCloserColumnSpan(physical[line] ?? '', closerKind),
            ));
            return;
        }
        // Anything above the matched opener was never closed.
        for (let k = stack.length - 1; k > idx; k--) {
            const open = stack[k];
            problems.push(fullLineProblem(
                physical, open.line,
                `Missing '${CLOSE_PHRASE[open.kind]}' for '${open.label}'.`,
                'error',
                {
                    code: 'missing-block-closer',
                    expectedClose: CLOSE_PHRASE[open.kind],
                    insertLine: line,
                },
                blockOpenerColumnSpan(physical[open.line] ?? '', open.kind),
            ));
        }
        stack.length = idx;
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }
        if (options.isInactiveLine?.(ll.line) && !isPreprocessorLine(t)) {
            continue;
        }

        if (/^#\s*If\b/i.test(t)) {
            preprocessorStack.push(ll.line);
            continue;
        }

        const preprocessorBranch = /^(#\s*ElseIf|#\s*Else)\b/i.exec(t);
        if (preprocessorBranch) {
            if (preprocessorStack.length === 0) {
                problems.push(fullLineProblem(
                    physical, ll.line,
                    `'${preprocessorBranch[1].replace(/\s+/g, ' ')}' has no matching '#If'.`,
                    'error',
                    { code: 'unmatched-block-closer' },
                    preprocessorBranchColumnSpan(physical[ll.line] ?? ''),
                ));
            }
            continue;
        }

        if (/^#\s*End\s*If\b/i.test(t) || /^#\s*EndIf\b/i.test(t)) {
            if (preprocessorStack.length === 0) {
                problems.push(fullLineProblem(
                    physical, ll.line,
                    `'#End If' has no matching '#If'.`,
                    'error',
                    { code: 'unmatched-block-closer' },
                    blockCloserColumnSpan(physical[ll.line] ?? '', 'PreprocessorIf'),
                ));
            } else {
                preprocessorStack.pop();
            }
            continue;
        }

        if (/^Next\b/i.test(t)) {
            const rest = t.replace(/^Next\b/i, '').trim();
            const count = rest === '' ? 1 : rest.split(',').length;
            for (let n = 0; n < count; n++) { closeOne('For', ll.line, 'Next'); }
            continue;
        }

        const closer = matchCloser(t);
        if (closer) {
            const word = closer === 'PreprocessorIf'
                ? '#End If'
                : /^End\b/i.test(t)
                    ? `End ${closer === 'Select' ? 'Select' : closer}`
                    : t.split(/\s+/)[0];
            closeOne(closer, ll.line, word);
            continue;
        }

        const opener = matchOpener(t);
        const isConditionalAlternativeHeader = !!opener &&
            isConditionalAlternativeProcedureHeader(stack, opener, preprocessorStack.length);
        const activeProcedure = activeProcedureBlock(stack);
        if (activeProcedure) {
            const declaration = moduleDeclarationInProcedureHit(physical[ll.line] ?? '');
            const procedureIndent = leadingWhitespace(physical[activeProcedure.line] ?? '').length;
            if (declaration && !isConditionalAlternativeHeader && declaration.span.startCol > procedureIndent) {
                problems.push(fullLineProblem(
                    physical,
                    ll.line,
                    `${declaration.label} must appear in the module declarations section, not inside a procedure.`,
                    'error',
                    { code: 'module-declaration-in-procedure' },
                    declaration.span,
                ));
            }
        }

        if (opener) {
            if (isTypeFieldNamedTypeInsideType(stack, opener, t)) {
                continue;
            }
            if (isConditionalAlternativeHeader) {
                continue;
            }
            stack.push({ ...opener, line: ll.line });
        }
    }

    for (const open of stack) {
        problems.push(fullLineProblem(
            physical, open.line,
            `Missing '${CLOSE_PHRASE[open.kind]}' for '${open.label}'.`,
            'error',
            {
                code: 'missing-block-closer',
                expectedClose: CLOSE_PHRASE[open.kind],
                insertLine: physical.length,
            },
            blockOpenerColumnSpan(physical[open.line] ?? '', open.kind),
        ));
    }
    for (const line of preprocessorStack) {
        problems.push(fullLineProblem(
            physical, line,
            `Missing '${CLOSE_PHRASE.PreprocessorIf}' for '#If'.`,
            'error',
            {
                code: 'missing-block-closer',
                expectedClose: CLOSE_PHRASE.PreprocessorIf,
                insertLine: physical.length,
            },
            blockOpenerColumnSpan(physical[line] ?? '', 'PreprocessorIf'),
        ));
    }

    problems.sort((a, b) => a.line - b.line || a.startCol - b.startCol);
    return problems;
}
