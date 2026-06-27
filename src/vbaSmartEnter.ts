// Pure, dependency-free smart-editing helpers for the vba language: Smart
// Enter block completion, With-member continuation, loop-iterator sync, and
// procedure-header paren repair. Keeping this module free of any `vscode`
// import means it can be unit-tested directly with vitest.
//
// The stripped-line substrate (strings/comments blanked, columns preserved)
// is derived from the analyzer lexer (src/analyzer/lexer/strippedLines.ts,
// audit #74), so one lexer defines string/comment semantics for Smart Enter
// and keyword completion. tests/smartEnterSubstrateComparison.test.ts pins
// the substrate against the legacy stripVba scanner corpus-wide.

import { lexerStrippedLine, lexerStrippedLines } from './analyzer/lexer/strippedLines';
import {
    leadingWhitespace,
    lineStartOffsets,
    logicalLinesFromStripped,
    VBA_IDENTIFIER_PATTERN,
} from './vbaSourceScan';
import { CLOSE_PHRASE, matchCloser, matchOpener, type BlockKind } from './vbaStructuralDiagnostics';

export interface VbaSmartBlockOpener {
    /** Closing statement smart-enter should insert. */
    endKeyword: string;
    /** Optional text to place on the indented body line before moving the caret. */
    bodyPrefix?: string;
}

export interface VbaSmartBlockInsertion {
    /** Text that should occupy the editable body line after Enter. */
    bodyText: string;
    /** Number of lines after the editor-created line where `bodyText` lands. */
    bodyLineOffset: number;
    /** Replacement for the editor-created body line. */
    replacementText: string;
}

export type VbaSmartBlockLayout = 'comfy' | 'compact';

export const VBA_BLOCK_INDENT_UNIT = '\t';
export const DEFAULT_VBA_SMART_BLOCK_LAYOUT: VbaSmartBlockLayout = 'comfy';

export function normalizeSmartBlockLayout(value: unknown): VbaSmartBlockLayout {
    return value === 'compact' ? 'compact' : DEFAULT_VBA_SMART_BLOCK_LAYOUT;
}

export interface VbaLoopIteratorSyncEdit {
    /** Absolute source span to replace. */
    span: { start: number; end: number };
    /** Iterator text copied from the edited side of the loop pair. */
    newText: string;
}

export interface VbaProcedureHeaderParensEdit {
    startCol: number;
    endCol: number;
    newText: string;
}

interface SmartOpenBlock {
    kind: BlockKind;
    closer: string;
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

export function procedureHeaderParensEdit(line: string): VbaProcedureHeaderParensEdit | undefined {
    const stripped = lexerStrippedLine(line);
    const patterns = [
        new RegExp(
            `^(\\s*(?:(?:Public|Private|Friend|Global)\\s+)?(?:Static\\s+)?Sub\\s+${VBA_IDENTIFIER_PATTERN})(\\s*)$`,
            'i',
        ),
        new RegExp(
            `^(\\s*(?:(?:Public|Private|Friend|Global)\\s+)?(?:Static\\s+)?Function\\s+${VBA_IDENTIFIER_PATTERN})(\\s*(?:As\\s+.+)?\\s*)$`,
            'i',
        ),
        new RegExp(
            `^(\\s*(?:(?:Public|Private|Friend|Global)\\s+)?(?:Static\\s+)?Property\\s+Get\\s+${VBA_IDENTIFIER_PATTERN})(\\s*(?:As\\s+.+)?\\s*)$`,
            'i',
        ),
    ];
    for (const pattern of patterns) {
        const match = pattern.exec(stripped);
        if (!match) {
            continue;
        }
        const col = match[1].length;
        return { startCol: col, endCol: col, newText: '()' };
    }
    return undefined;
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
    const logical = logicalLinesFromStripped(lexerStrippedLines(source.slice(0, safeOffset)));
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
    const safeOffset = Math.max(0, Math.min(offset, source.length));
    // Cheap early-out: most edits are not on a For/Next iterator, so parse
    // just the edited physical line before splitting the whole document.
    const info = parseLoopLine(physicalLineAt(source, safeOffset));
    if (!info?.iterator || !offsetTouchesSpan(safeOffset, info.iterator.span)) {
        return undefined;
    }

    const lines = physicalLines(source);
    const lineIndex = physicalLineAtOffset(lines, safeOffset);

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
 * Builds the exact body-line replacement for Smart Enter. The normal shape is:
 * opener line, spacer line, one indented editable body line, spacer line, then
 * the closer on its own line.
 */
export function smartBlockInsertion(
    openerLine: string,
    currentBodyLine: string,
    opener: VbaSmartBlockOpener,
    options: {
        eol?: string;
        insertCloser?: boolean;
        indentUnit?: string;
        layout?: VbaSmartBlockLayout;
    } = {},
): VbaSmartBlockInsertion {
    const bodyText = smartBlockBodyText(
        openerLine,
        currentBodyLine,
        opener,
        options.indentUnit,
    );
    const eol = options.eol ?? '\n';
    if (options.insertCloser === false) {
        return {
            bodyText,
            bodyLineOffset: 0,
            replacementText: bodyText,
        };
    }
    const layout = options.layout ?? DEFAULT_VBA_SMART_BLOCK_LAYOUT;
    if (layout === 'compact') {
        return {
            bodyText,
            bodyLineOffset: 0,
            replacementText: `${bodyText}${eol}${leadingWhitespace(openerLine)}${opener.endKeyword}`,
        };
    }
    return {
        bodyText,
        bodyLineOffset: 1,
        replacementText: `${eol}${bodyText}${eol}${eol}${leadingWhitespace(openerLine)}${opener.endKeyword}`,
    };
}

/**
 * When Enter is pressed after a leading-dot member line inside an active With
 * block, the next line should keep the same indentation and seed another dot.
 */
export function withMemberContinuationText(
    source: string,
    previousLineIndex: number,
): string | undefined {
    const lines = source.split(/\r\n|\r|\n/);
    const previousLine = lines[previousLineIndex];
    if (previousLine === undefined) {
        return undefined;
    }
    if (!/^[ \t]*\./.test(lexerStrippedLine(previousLine))) {
        return undefined;
    }

    const starts = lineStartOffsets(source);
    const previousLineEnd = (starts[previousLineIndex] ?? 0) + previousLine.length;
    if (!openSmartBlockClosersBefore(source, previousLineEnd).includes('End With')) {
        return undefined;
    }

    return `${leadingWhitespace(previousLine)}.`;
}

/**
 * When Enter is pressed at the end of a whole-line VBA comment, the new line
 * continues the comment: the same indentation, the same apostrophe run, and
 * (when mirrorSpacing is on) the same run of spaces that followed the apostrophe
 * so the comment text lines up. Returns undefined when the line above is not a
 * whole-line comment (a trailing comment after code never continues).
 */
export function commentContinuationText(
    source: string,
    previousLineIndex: number,
    mirrorSpacing: boolean,
): string | undefined {
    const lines = source.split(/\r\n|\r|\n/);
    const previousLine = lines[previousLineIndex];
    if (previousLine === undefined) {
        return undefined;
    }
    // Leading indentation, the apostrophe run, then the spaces that follow it.
    const match = /^([ \t]*)('+)([ \t]*)/.exec(previousLine);
    if (!match) {
        return undefined;
    }
    const [, indent, apostrophes, spacesAfter] = match;
    return `${indent}${apostrophes}${mirrorSpacing ? spacesAfter : ''}`;
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

// Extracts the single physical line containing `offset` without splitting the
// whole document; mirrors physicalLineAtOffset, including its fall-back to the
// final line when the offset sits on the LF of a CRLF pair.
function physicalLineAt(source: string, offset: number): PhysicalLine {
    const start = offset === 0 ? 0 : source.lastIndexOf('\n', offset - 1) + 1;
    const nl = source.indexOf('\n', offset);
    const end = nl === -1
        ? source.length
        : (nl > start && source[nl - 1] === '\r' ? nl - 1 : nl);
    if (offset <= end) {
        return { text: source.slice(start, end), start, end };
    }
    const lastStart = source.lastIndexOf('\n') + 1;
    return { text: source.slice(lastStart), start: lastStart, end: source.length };
}

function parseLoopLine(line: PhysicalLine): LoopLineInfo | undefined {
    const stripped = lexerStrippedLine(line.text);

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
