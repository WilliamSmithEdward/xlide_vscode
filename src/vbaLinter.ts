// Pure, dependency-free VBA structural analysis used by the diagnostics
// provider and the smart "auto End Sub" editing feature. Keeping this module
// free of any `vscode` import means it can be unit-tested directly with vitest.

export interface VbaLintProblem {
    /** 0-based physical line of the relevant token. */
    line: number;
    /** 0-based start column (inclusive). */
    startCol: number;
    /** 0-based end column (exclusive). */
    endCol: number;
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
    | 'For' | 'Do' | 'While';

interface OpenBlock {
    kind: BlockKind;
    /** 0-based physical line of the opener. */
    line: number;
    /** Friendly descriptor, e.g. "Sub Foo" or "If". */
    label: string;
}

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
): VbaLintProblem {
    const raw = physical[line] ?? '';
    const startCol = raw.length - raw.trimStart().length;
    return { line, startCol, endCol: Math.max(raw.length, startCol + 1), message, severity };
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
            ));
        }
        stack.length = idx;
    };

    for (const ll of logical) {
        const t = ll.text.trim();
        if (!t) { continue; }

        if (/^Next\b/i.test(t)) {
            const rest = t.replace(/^Next\b/i, '').trim();
            const count = rest === '' ? 1 : rest.split(',').length;
            for (let n = 0; n < count; n++) { closeOne('For', ll.line, 'Next'); }
            continue;
        }

        const closer = matchCloser(t);
        if (closer) {
            const word = /^End\b/i.test(t) ? `End ${closer === 'Select' ? 'Select' : closer}` : t.split(/\s+/)[0];
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
        ));
    }

    problems.sort((a, b) => a.line - b.line || a.startCol - b.startCol);
    return problems;
}

/**
 * If the stripped line opens a Sub/Function/Property, returns the matching
 * `End` keyword to auto-insert. Used by the smart-enter editing feature.
 */
export function detectProcOpener(strippedLine: string): { endKeyword: string } | undefined {
    const t = strippedLine.trim();
    if (/^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Sub\s+[A-Za-z_]\w*/i.test(t)) {
        return { endKeyword: 'End Sub' };
    }
    if (/^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Function\s+[A-Za-z_]\w*/i.test(t)) {
        return { endKeyword: 'End Function' };
    }
    if (/^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?Property\s+(?:Get|Let|Set)\s+[A-Za-z_]\w*/i.test(t)) {
        return { endKeyword: 'End Property' };
    }
    return undefined;
}

/**
 * Returns true if the procedure that opens at `openerIdx` already has its
 * matching `endKeyword` before the next procedure opener or end of file.
 * `strippedLines` must already have strings/comments removed.
 */
export function isProcClosedAhead(
    strippedLines: string[], openerIdx: number, endKeyword: string,
): boolean {
    const endRe = new RegExp('^' + endKeyword.replace(' ', '\\s+') + '\\b', 'i');
    const otherOpenerRe =
        /^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(?:Sub|Function|Property)\b/i;
    for (let i = openerIdx + 1; i < strippedLines.length; i++) {
        const t = strippedLines[i].trim();
        if (!t) { continue; }
        if (endRe.test(t)) { return true; }
        if (otherOpenerRe.test(t)) { return false; }
    }
    return false;
}
