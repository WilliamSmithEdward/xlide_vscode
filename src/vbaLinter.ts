// Pure, dependency-free VBA structural analysis used by the diagnostics
// provider and the smart auto-block editing feature. Keeping this module
// free of any `vscode` import means it can be unit-tested directly with vitest.

export interface VbaLintProblem {
    /** 0-based physical line of the relevant token. */
    line: number;
    /** 0-based start column (inclusive). */
    startCol: number;
    /** 0-based end column (exclusive). */
    endCol: number;
    /** Stable diagnostic code for editor integrations. */
    code?: 'missing-block-closer' | 'unmatched-block-closer';
    /** Closing phrase that can repair a missing-block diagnostic. */
    expectedClose?: string;
    /** 0-based physical line before which the missing closer should be inserted. */
    insertLine?: number;
    message: string;
    severity: 'error' | 'warning';
}

/** A logical line after string/comment stripping and continuation joining. */
interface LogicalLine {
    /** Stripped, continuation-joined text. */
    text: string;
    /** 0-based physical line where this logical line begins. */
    line: number;
}

type BlockKind =
    | 'Sub' | 'Function' | 'Property'
    | 'If' | 'With' | 'Select' | 'Type' | 'Enum'
    | 'For' | 'Do' | 'While' | 'PreprocessorIf';

export interface VbaSmartBlockOpener {
    /** Closing statement smart-enter should insert. */
    endKeyword: string;
    /** Optional text to place on the indented body line before moving the caret. */
    bodyPrefix?: string;
}

export const VBA_BLOCK_INDENT_UNIT = '\t';

export interface VbaLoopIteratorSyncEdit {
    /** Absolute source span to replace. */
    span: { start: number; end: number };
    /** Iterator text copied from the edited side of the loop pair. */
    newText: string;
}

interface SmartOpenBlock {
    kind: BlockKind;
    closer: string;
}

interface OpenBlock {
    kind: BlockKind;
    /** 0-based physical line of the opener. */
    line: number;
    /** Friendly descriptor, e.g. "Sub Foo" or "If". */
    label: string;
}

interface PhysicalLine {
    text: string;
    start: number;
    end: number;
}

interface LoopIteratorToken {
    name: string;
    span: { start: number; end: number };
}

interface LoopOpenerLine {
    kind: 'opener';
    iterator: LoopIteratorToken;
}

interface LoopNextLine {
    kind: 'next';
    closeCount: number;
    iterator?: LoopIteratorToken;
}

type LoopLineInfo = LoopOpenerLine | LoopNextLine;

/** The closing phrase expected for each block kind. */
const CLOSE_PHRASE: Record<BlockKind, string> = {
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

/**
 * Replaces string-literal contents and trailing comments with spaces so that
 * keyword detection never trips over text inside quotes or comments, while
 * keeping every column aligned with the original line.
 */
export function stripVba(line: string): string {
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
                } else {
                    chars[i] = ' ';
                    inString = false;
                }
            } else {
                chars[i] = ' ';
            }
        } else if (c === '"') {
            chars[i] = ' ';
            inString = true;
        } else if (c === "'") {
            for (let j = i; j < chars.length; j++) { chars[j] = ' '; }
            break;
        }
    }
    let out = chars.join('');
    // Blank out a `Rem` comment (whole-line form) keeping column alignment.
    const rem = /^(\s*)Rem\b/i.exec(out);
    if (rem) {
        out = out.slice(0, rem[1].length) + ' '.repeat(out.length - rem[1].length);
    }
    return out;
}

/** Splits source, strips each line, and joins `_` line continuations. */
function toLogicalLines(source: string): { stripped: string[]; logical: LogicalLine[] } {
    const physical = source.split(/\r\n|\r|\n/);
    const stripped = physical.map(stripVba);
    const logical: LogicalLine[] = [];
    let i = 0;
    while (i < stripped.length) {
        let text = stripped[i];
        const startLine = i;
        while (/\s_[ \t]*$/.test(text) && i + 1 < stripped.length) {
            text = text.replace(/\s_[ \t]*$/, ' ') + stripped[i + 1];
            i++;
        }
        logical.push({ text, line: startLine });
        i++;
    }
    return { stripped, logical };
}

/** Detects a block closer on a stripped, trimmed logical line. */
function matchCloser(t: string): BlockKind | undefined {
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
function matchOpener(t: string): OpenBlock | undefined {
    let m: RegExpExecArray | null;
    if (/^#\s*If\b/i.test(t)) { return { kind: 'PreprocessorIf', line: 0, label: '#If' }; }
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) {
        const kind = (/^sub$/i.test(m[1]) ? 'Sub' : 'Function') as BlockKind;
        return { kind, line: 0, label: `${kind} ${m[2]}` };
    }
    m = /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) { return { kind: 'Property', line: 0, label: `Property ${m[2]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Type\s+([A-Za-z_]\w*)/i.exec(t);
    if (m) { return { kind: 'Type', line: 0, label: `Type ${m[1]}` }; }
    m = /^(?:(?:Public|Private|Global)\s+)?Enum\s+([A-Za-z_]\w*)/i.exec(t);
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
    details: Pick<VbaLintProblem, 'code' | 'expectedClose' | 'insertLine'> = {},
): VbaLintProblem {
    const raw = physical[line] ?? '';
    const startCol = raw.length - raw.trimStart().length;
    return {
        line,
        startCol,
        endCol: Math.max(raw.length, startCol + 1),
        message,
        severity,
        ...details,
    };
}

/**
 * Performs structural block-balance analysis on VBA source, reporting:
 *  - openers with no matching closer (missing `End Sub`, `Next`, ...),
 *  - closers with no matching opener (stray `End If`, `Loop`, ...),
 *  - mismatched nesting (an inner block left unclosed).
 */
export function lintVbaSource(source: string): VbaLintProblem[] {
    const physical = source.split(/\r\n|\r|\n/);
    const { logical } = toLogicalLines(source);
    const stack: OpenBlock[] = [];
    const problems: VbaLintProblem[] = [];

    const closeOne = (closerKind: BlockKind, line: number, closerWord: string): void => {
        let idx = -1;
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].kind === closerKind) { idx = k; break; }
        }
        if (idx === -1) {
            problems.push(fullLineProblem(
                physical, line,
                `'${closerWord}' has no matching '${OPEN_WORD[closerKind]}'.`,
                'error',
                { code: 'unmatched-block-closer' },
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
            ));
        }
        stack.length = idx;
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }

        const preprocessorBranch = /^(#\s*ElseIf|#\s*Else)\b/i.exec(t);
        if (preprocessorBranch) {
            if (!stack.some((b) => b.kind === 'PreprocessorIf')) {
                problems.push(fullLineProblem(
                    physical, ll.line,
                    `'${preprocessorBranch[1].replace(/\s+/g, ' ')}' has no matching '#If'.`,
                    'error',
                    { code: 'unmatched-block-closer' },
                ));
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
        if (opener) {
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
        ));
    }

    problems.sort((a, b) => a.line - b.line || a.startCol - b.startCol);
    return problems;
}

/**
 * If the stripped line opens a block that is safe to complete immediately after
 * Enter, returns the closing statement and any body-line prefix.
 */
export function detectSmartBlockOpener(strippedLine: string): VbaSmartBlockOpener | undefined {
    const t = strippedLine.trim();
    if (!t || t.includes(':')) {
        return undefined;
    }

    const opener = matchOpener(t);
    if (!opener) {
        return undefined;
    }
    if (!isCompleteSmartBlockOpener(t, opener.kind)) {
        return undefined;
    }

    if (opener.kind === 'For') {
        const iterator = forIteratorName(t);
        return { endKeyword: iterator ? `Next ${iterator}` : 'Next' };
    }
    if (opener.kind === 'With') {
        return { endKeyword: CLOSE_PHRASE[opener.kind], bodyPrefix: '.' };
    }
    return { endKeyword: CLOSE_PHRASE[opener.kind] };
}

/**
 * Returns the still-open smart-block closers before an offset, ordered from
 * outermost to innermost. Used by keyword completion so close suggestions and
 * Smart Enter share the same block-opening rules.
 */
export function openSmartBlockClosersBefore(
    source: string,
    offset = source.length,
): string[] {
    const safeOffset = Math.max(0, Math.min(offset, source.length));
    const { logical } = toLogicalLines(source.slice(0, safeOffset));
    const stack: SmartOpenBlock[] = [];

    const closeOne = (kind: BlockKind): void => {
        let idx = -1;
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].kind === kind) {
                idx = k;
                break;
            }
        }
        if (idx >= 0) {
            stack.length = idx;
        }
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }
        if (/^(#\s*ElseIf|#\s*Else)\b/i.test(t)) { continue; }

        if (/^Next\b/i.test(t)) {
            const rest = t.replace(/^Next\b/i, '').trim();
            const count = rest === '' ? 1 : rest.split(',').length;
            for (let n = 0; n < count; n++) { closeOne('For'); }
            continue;
        }

        const closer = matchCloser(t);
        if (closer) {
            closeOne(closer);
            continue;
        }

        const opener = detectSmartBlockOpener(t);
        const matched = opener ? matchOpener(t) : undefined;
        if (opener && matched) {
            stack.push({ kind: matched.kind, closer: opener.endKeyword });
        }
    }

    return stack.map((open) => open.closer);
}

/**
 * When the edit position is on a simple `For` / `For Each` iterator or its
 * matching `Next name`, returns the paired iterator replacement.
 */
export function resolveLoopIteratorSyncEdit(
    source: string,
    offset: number,
): VbaLoopIteratorSyncEdit | undefined {
    const lines = physicalLines(source);
    if (lines.length === 0) { return undefined; }

    const safeOffset = Math.max(0, Math.min(offset, source.length));
    const lineIndex = physicalLineAtOffset(lines, safeOffset);
    const line = lines[lineIndex];
    const info = parseLoopLine(line);
    if (!info?.iterator || !offsetTouchesSpan(safeOffset, info.iterator.span)) {
        return undefined;
    }

    const counterpart = info.kind === 'opener'
        ? findMatchingNextLine(lines, lineIndex)
        : findMatchingOpenerLine(lines, lineIndex);
    if (!counterpart?.iterator) {
        return undefined;
    }
    if (counterpart.iterator.name === info.iterator.name) {
        return undefined;
    }

    return {
        span: counterpart.iterator.span,
        newText: info.iterator.name,
    };
}

/**
 * Computes the body-line indentation Smart Enter should own after a block
 * opener. If the editor already produced a deeper indent, keep it; otherwise
 * move one configured indentation unit deeper than the opener.
 */
export function smartBlockBodyIndent(
    openerLine: string,
    currentBodyLine = '',
    indentUnit = VBA_BLOCK_INDENT_UNIT,
): string {
    const openerIndent = leadingWhitespace(openerLine);
    const currentIndent = leadingWhitespace(currentBodyLine);
    const expectedIndent = openerIndent + indentUnit;
    if (currentIndent.startsWith(expectedIndent)) {
        return currentIndent;
    }
    return expectedIndent;
}

/**
 * Body-line text Smart Enter should leave after a block opener. For `With`,
 * this includes the seeded leading dot after the computed indentation.
 */
export function smartBlockBodyText(
    openerLine: string,
    currentBodyLine: string,
    opener: Pick<VbaSmartBlockOpener, 'bodyPrefix'>,
    indentUnit = VBA_BLOCK_INDENT_UNIT,
): string {
    return smartBlockBodyIndent(openerLine, currentBodyLine, indentUnit) + (opener.bodyPrefix ?? '');
}

/**
 * Back-compat wrapper for tests/providers that only care about procedure
 * headers. New smart-enter callers should use `detectSmartBlockOpener`.
 */
export function detectProcOpener(strippedLine: string): { endKeyword: string } | undefined {
    const opener = detectSmartBlockOpener(strippedLine);
    if (
        opener?.endKeyword === 'End Sub' ||
        opener?.endKeyword === 'End Function' ||
        opener?.endKeyword === 'End Property'
    ) {
        return { endKeyword: opener.endKeyword };
    }
    return undefined;
}

/**
 * Returns true if the smart-enter block opener already has a compatible closer
 * before the next procedure opener or end of file.
 * `strippedLines` must already have strings/comments removed.
 */
export function isSmartBlockClosedAhead(
    strippedLines: string[], openerIdx: number, opener: VbaSmartBlockOpener,
): boolean {
    const endRe = closerRegex(opener.endKeyword);
    for (let i = openerIdx + 1; i < strippedLines.length; i++) {
        const t = strippedLines[i].trim();
        if (!t) { continue; }
        if (endRe.test(t)) { return true; }
        if (isProcedureOpener(t)) { return false; }
    }
    return false;
}

/**
 * Back-compat wrapper for procedure-only smart-enter callers.
 */
export function isProcClosedAhead(
    strippedLines: string[], openerIdx: number, endKeyword: string,
): boolean {
    return isSmartBlockClosedAhead(strippedLines, openerIdx, { endKeyword });
}

function forIteratorName(t: string): string | undefined {
    return /^For\s+(?:Each\s+)?([A-Za-z_]\w*)\b/i.exec(t)?.[1];
}

function physicalLines(source: string): PhysicalLine[] {
    const lines: PhysicalLine[] = [];
    let start = 0;
    for (let i = 0; i < source.length; i++) {
        if (source[i] !== '\n') { continue; }
        const end = i > start && source[i - 1] === '\r' ? i - 1 : i;
        lines.push({ text: source.slice(start, end), start, end });
        start = i + 1;
    }
    lines.push({ text: source.slice(start), start, end: source.length });
    return lines;
}

function physicalLineAtOffset(lines: PhysicalLine[], offset: number): number {
    for (let i = 0; i < lines.length; i++) {
        if (offset >= lines[i].start && offset <= lines[i].end) {
            return i;
        }
    }
    return lines.length - 1;
}

function parseLoopLine(line: PhysicalLine): LoopLineInfo | undefined {
    const stripped = stripVba(line.text);

    let m = /^(\s*)For\s+Each\s+([A-Za-z_]\w*)\s+In\b/i.exec(stripped);
    if (m) {
        return { kind: 'opener', iterator: tokenFromMatch(line.start, m, 2) };
    }

    m = /^(\s*)For\s+([A-Za-z_]\w*)\s*=/i.exec(stripped);
    if (m) {
        return { kind: 'opener', iterator: tokenFromMatch(line.start, m, 2) };
    }

    const next = /^(\s*)Next\b(.*)$/i.exec(stripped);
    if (!next) {
        return undefined;
    }

    const closeCount = nextCloseCount(next[2].trim());
    m = /^(\s*)Next\s+([A-Za-z_]\w*)\s*$/i.exec(stripped);
    return {
        kind: 'next',
        closeCount,
        iterator: m ? tokenFromMatch(line.start, m, 2) : undefined,
    };
}

function tokenFromMatch(lineStart: number, match: RegExpExecArray, groupIndex: number): LoopIteratorToken {
    const name = match[groupIndex];
    const col = match[0].indexOf(name);
    return {
        name,
        span: { start: lineStart + col, end: lineStart + col + name.length },
    };
}

function nextCloseCount(rest: string): number {
    if (!rest) { return 1; }
    return Math.max(1, rest.split(',').map((part) => part.trim()).filter(Boolean).length);
}

function findMatchingNextLine(lines: PhysicalLine[], openerIndex: number): LoopNextLine | undefined {
    let depth = 0;
    for (let i = openerIndex + 1; i < lines.length; i++) {
        const info = parseLoopLine(lines[i]);
        if (!info) { continue; }
        if (info.kind === 'opener') {
            depth++;
            continue;
        }

        if (depth === 0) {
            return info.closeCount === 1 ? info : undefined;
        }

        depth -= info.closeCount;
        if (depth < 0) {
            return undefined;
        }
    }
    return undefined;
}

function findMatchingOpenerLine(lines: PhysicalLine[], nextIndex: number): LoopOpenerLine | undefined {
    let depth = 0;
    for (let i = nextIndex - 1; i >= 0; i--) {
        const info = parseLoopLine(lines[i]);
        if (!info) { continue; }
        if (info.kind === 'next') {
            depth += info.closeCount;
            continue;
        }

        if (depth === 0) {
            return info;
        }
        depth--;
    }
    return undefined;
}

function offsetTouchesSpan(offset: number, span: { start: number; end: number }): boolean {
    return offset >= span.start && offset <= span.end;
}

function leadingWhitespace(text: string): string {
    return /^[ \t]*/.exec(text)?.[0] ?? '';
}

function isCompleteSmartBlockOpener(t: string, kind: BlockKind): boolean {
    switch (kind) {
        case 'With':
            return /^With\s+\S/i.test(t);
        case 'If':
            return /^If\s+\S.+\bThen\s*$/i.test(t);
        case 'For':
            return (
                /^For\s+Each\s+[A-Za-z_]\w*\s+In\s+\S/i.test(t) ||
                /^For\s+[A-Za-z_]\w*\s*=.+\bTo\b.+/i.test(t)
            );
        case 'Do':
            return /^Do\s*$/i.test(t) || /^Do\s+(?:While|Until)\s+\S/i.test(t);
        case 'Select':
            return /^Select\s+Case\s+\S/i.test(t);
        case 'While':
            return /^While\s+\S/i.test(t);
        case 'PreprocessorIf':
            return /^#\s*If\s+.+\bThen\s*$/i.test(t);
        default:
            return true;
    }
}

function isProcedureOpener(t: string): boolean {
    return /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property)\b/i.test(t);
}

function closerRegex(endKeyword: string): RegExp {
    if (/^Next\b/i.test(endKeyword)) {
        return /^Next\b/i;
    }
    const escaped = endKeyword
        .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        .replace(/\s+/g, '\\s+');
    return new RegExp(`^${escaped}\\b`, 'i');
}
