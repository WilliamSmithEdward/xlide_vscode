// Pure, dependency-free VBA source-scan utilities shared by the structural
// analyzer, smart editing features, rename, the test runner, and commands.
// Keeping this module free of any `vscode` import means it can be unit-tested
// directly with vitest.

import { isReservedIdentifier } from './analyzer/lexer/keywordTable';

export const VBA_IDENTIFIER_PATTERN = '[A-Za-z_][A-Za-z0-9_]*';
export const VBA_IDENTIFIER_RE = /[A-Za-z_][A-Za-z0-9_]*/;
export const VBA_IDENTIFIER_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
export const VBA_MODULE_NAME_MAX_LENGTH = 31;

/** Input-box validator for VBA module names: undefined when valid, message otherwise. */
export function validateVbaModuleName(name: string): string | undefined {
    if (!VBA_IDENTIFIER_NAME_RE.test(name)) {
        return 'Module names must start with a letter or underscore and use only letters, digits, and underscores';
    }
    if (name.length > VBA_MODULE_NAME_MAX_LENGTH) {
        return `Module names must be at most ${VBA_MODULE_NAME_MAX_LENGTH} characters`;
    }
    if (isReservedIdentifier(name)) {
        return 'Module names cannot be VBA reserved words';
    }
    return undefined;
}

const VBA_IDENTIFIER_WORD_RE = /[A-Za-z_][A-Za-z0-9_]*/g;

export interface VbaIdentifierOccurrence {
    line: number;
    column: number;
    offset: number;
    text: string;
}

/** A logical line after string/comment stripping and continuation joining. */
export interface LogicalLine {
    /** Stripped, continuation-joined text. */
    text: string;
    /** 0-based physical line where this logical line begins. */
    line: number;
}

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
    // Blank out a `Rem` comment to end of line, keeping column alignment. A `Rem`
    // comment begins at a statement start: the line start OR after a `:` statement
    // separator (e.g. `x = 1: Rem note`). Blanking only the whole-line form let a
    // `: Rem ...` comment's text - including any `:` inside it - leak into the
    // colon-split logical lines as phantom block openers/closers.
    const rem = /(^|:)([ \t]*)Rem\b/i.exec(out);
    if (rem) {
        const remKeywordStart = rem.index + rem[1].length + rem[2].length;
        out = out.slice(0, remKeywordStart) + ' '.repeat(out.length - remKeywordStart);
    }
    return out;
}

/**
 * Precomputes the absolute offset at which each physical line starts.
 * Treats CR, CRLF, and LF as line terminators, matching the analyzer lexer
 * and VS Code's own line splitting.
 */
export function lineStartOffsets(source: string): number[] {
    const starts = [0];
    for (let i = 0; i < source.length; i++) {
        const ch = source[i];
        if (ch === '\r') {
            if (source[i + 1] === '\n') {
                i++;
            }
            starts.push(i + 1);
        } else if (ch === '\n') {
            starts.push(i + 1);
        }
    }
    return starts;
}

/** Returns the leading spaces/tabs for a source line or snippet of text. */
export function leadingWhitespace(text: string): string {
    return /^[ \t]*/.exec(text)?.[0] ?? '';
}

/** Normalizes CR, CRLF, and LF line terminators to LF. */
export function normalizeEol(text: string): string {
    return text.replace(/\r\n|\r/g, '\n');
}

/** Line terminator to use when re-serializing edits to `source`. */
export function detectEol(source: string): string {
    return source.includes('\r\n') ? '\r\n' : '\n';
}

/** True for VBE attribute header lines such as `Attribute VB_Name = "..."`. */
export function isVbaAttributeLine(line: string): boolean {
    return /^\s*Attribute\s+(?:VB_[A-Za-z0-9_]+|[A-Za-z_][\w.]*\.VB_[A-Za-z0-9_]+)\s*=/.test(line);
}

/**
 * Finds whole-word identifier occurrences while ignoring strings and comments.
 * Offsets are absolute source offsets so callers do not recompute line starts.
 */
export function findIdentifierOccurrences(
    source: string,
    name: string,
): VbaIdentifierOccurrence[] {
    const lines = source.split(/\r\n|\r|\n/);
    const starts = lineStartOffsets(source);
    const lower = name.toLowerCase();
    const out: VbaIdentifierOccurrence[] = [];
    for (let i = 0; i < lines.length; i++) {
        const stripped = stripVba(lines[i]);
        VBA_IDENTIFIER_WORD_RE.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = VBA_IDENTIFIER_WORD_RE.exec(stripped)) !== null) {
            if (m[0].toLowerCase() === lower) {
                out.push({
                    line: i,
                    column: m.index,
                    offset: (starts[i] ?? 0) + m.index,
                    text: m[0],
                });
            }
        }
    }
    return out;
}

/** Splits source, strips each line, and joins `_` line continuations. */
export function toLogicalLines(source: string): { stripped: string[]; logical: LogicalLine[] } {
    const stripped = source.split(/\r\n|\r|\n/).map(stripVba);
    return { stripped, logical: logicalLinesFromStripped(stripped) };
}

/**
 * Joins `_` line continuations and splits `:` statement separators over lines
 * that already have strings/comments blanked. Shared by the legacy stripVba
 * substrate ({@link toLogicalLines}) and the analyzer-lexer substrate Smart
 * Enter sits on (src/analyzer/lexer/strippedLines.ts, audit #74).
 */
export function logicalLinesFromStripped(stripped: string[]): LogicalLine[] {
    const logical: LogicalLine[] = [];
    let i = 0;
    while (i < stripped.length) {
        let text = stripped[i];
        const startLine = i;
        while (/\s_[ \t]*$/.test(text) && i + 1 < stripped.length) {
            text = text.replace(/\s_[ \t]*$/, ' ') + stripped[i + 1];
            i++;
        }
        for (const segment of splitColonStatements(text)) {
            logical.push({ text: segment, line: startLine });
        }
        i++;
    }
    return logical;
}

function splitColonStatements(text: string): string[] {
    const out: string[] = [];
    let start = 0;
    for (let i = 0; i < text.length; i++) {
        if (text[i] !== ':') {
            continue;
        }
        out.push(`${' '.repeat(start)}${text.slice(start, i)}`);
        start = i + 1;
    }
    out.push(`${' '.repeat(start)}${text.slice(start)}`);
    return out.filter((segment) => segment.trim().length > 0);
}
