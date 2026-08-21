// Pure, dependency-free VBA source-scan utilities shared by the structural
// analyzer, smart editing features, rename, the test runner, and commands.
// Keeping this module free of any `vscode` import means it can be unit-tested
// directly with vitest.

import { isReservedIdentifier } from './analyzer/lexer/keywordTable';

// VBA identifiers may use any locale letter, and a combining mark continues a
// name (Thai and Devanagari build a letter from a base plus a mark). The
// ASCII-only forms made go-to-definition, find-references, rename and
// smart-enter block completion do nothing at all on a Cyrillic, Greek, Thai or
// Japanese name: VS Code selects the word with VBA_IDENTIFIER_RE, and it
// matched none of it.
//
// Any regex built from VBA_IDENTIFIER_PATTERN needs the `u` flag, or \p{L} is
// read as a literal 'p{L}' and silently matches nothing.
export const VBA_IDENTIFIER_PATTERN = '[\\p{L}_][\\p{L}\\p{M}\\p{N}_]*';
export const VBA_IDENTIFIER_RE = /[\p{L}_][\p{L}\p{M}\p{N}_]*/u;
export const VBA_IDENTIFIER_NAME_RE = /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u;
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

const VBA_IDENTIFIER_WORD_RE = /[\p{L}_][\p{L}\p{M}\p{N}_]*/gu;

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
 * Drops the block a class or UserForm header opens with, returning the index of
 * the first line after it.
 *
 * A standard module's header is attribute lines alone, but a class opens with
 * `VERSION 1.0 CLASS` / `BEGIN` / ... / `END` and a UserForm with `VERSION 5.00`
 * / `Begin {GUID} Name` / ... / `End`, the form's block nesting once per
 * control. Filtering attribute lines alone left all of that on screen, so
 * hiding headers did nothing on a .cls and even less on a .frm.
 *
 * A block that never closes is left alone: showing too much beats hiding code.
 */
export function vbaHeaderBlockEnd(lines: readonly string[]): number {
    let index = 0;
    while (index < lines.length && lines[index].trim() === '') {
        index += 1;
    }
    if (index >= lines.length || !/^\s*VERSION\b/i.test(lines[index])) {
        return 0;
    }
    index += 1;
    while (index < lines.length && lines[index].trim() === '') {
        index += 1;
    }
    if (index >= lines.length || !/^\s*Begin\b/i.test(lines[index])) {
        // `VERSION` with no block after it: drop just that line.
        return index;
    }
    let depth = 0;
    for (; index < lines.length; index += 1) {
        const line = lines[index];
        if (/^\s*Begin\b/i.test(line)) {
            depth += 1;
            continue;
        }
        if (/^\s*End\b/i.test(line)) {
            depth -= 1;
            if (depth === 0) {
                return index + 1;
            }
            continue;
        }
        // Inside the block only property assignments and blank lines are legal.
        // Anything else means the block never closed, and scanning on would let
        // a body `End Sub` close it and take the whole module with it.
        if (line.trim() !== '' && !/^\s*[\w.]+\s*=/.test(line)) {
            return 0;
        }
    }
    return 0;
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

/** True for a segment that is exactly an `If ... Then` header, opening a block. */
const BLOCK_IF_HEADER_RE = /^If\b.*\bThen$/i;

function splitColonStatements(text: string): string[] {
    const out: string[] = [];
    let start = 0;
    for (let i = 0; i < text.length; i++) {
        if (text[i] !== ':') {
            continue;
        }
        const segment = text.slice(start, i);
        // A colon straight after `Then` separates statements INSIDE a
        // single-line If; it does not end the header. Splitting there left a
        // bare `If ... Then` segment, which reads as a block opener, so
        // `If utc_NegativeOffset Then: utc_Offset = -utc_Offset` - a common
        // idiom, and twice in Tim Hall's VBA-JSON - reported a missing End If
        // against the enclosing block.
        if (BLOCK_IF_HEADER_RE.test(segment.trim())) {
            continue;
        }
        out.push(`${' '.repeat(start)}${segment}`);
        start = i + 1;
    }
    out.push(`${' '.repeat(start)}${text.slice(start)}`);
    return out.filter((segment) => segment.trim().length > 0);
}
