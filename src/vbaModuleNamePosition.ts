// Recognising a module name in code.
//
// Issue #9 rule 3: a module's name is not a symbol - no declaration in any
// module's text holds it, it belongs to the component. So a word is a module
// reference only when BOTH hold: nothing else resolves at that position, and it
// stands somewhere only a module can stand.
//
// Asking the name alone is a bug worth not repeating: a workbook with modules
// called Log, Config or Data also has LOCALS called Log, Config and Data, and
// treating the word as a module renamed the module and left the variable alone.
//
// This file answers only the second half - the position test. The caller
// supplies the first half by asking it just once nothing else resolved.

/** Why a position can only be a module name. */
export type ModuleNamePositionKind = 'qualifier' | 'as' | 'new' | 'implements';

/**
 * Returns why the identifier occupying [start, end) can only be a module name,
 * or undefined when the position does not require one.
 */
export function moduleNamePositionKind(
    source: string,
    start: number,
    end: number,
): ModuleNamePositionKind | undefined {
    if (start < 0 || end > source.length || start >= end) {
        return undefined;
    }
    // Before a dot: `Helpers.Recalculate`. A line continuation counts as
    // whitespace, so `Helpers _` / `.Recalculate` is one qualified reference.
    if (followedByMemberDot(source, end)) {
        return 'qualifier';
    }
    const keyword = precedingKeyword(source, start);
    if (keyword === 'as' || keyword === 'new' || keyword === 'implements') {
        return keyword;
    }
    return undefined;
}

function followedByMemberDot(source: string, from: number): boolean {
    let i = from;
    for (;;) {
        while (i < source.length && (source[i] === ' ' || source[i] === '\t')) {
            i += 1;
        }
        if (source[i] === '_' && isLineBreakAfterContinuation(source, i + 1)) {
            i = skipToNextLine(source, i + 1);
            continue;
        }
        return source[i] === '.';
    }
}

function isLineBreakAfterContinuation(source: string, from: number): boolean {
    let i = from;
    while (i < source.length && (source[i] === ' ' || source[i] === '\t')) {
        i += 1;
    }
    return source[i] === '\r' || source[i] === '\n';
}

function skipToNextLine(source: string, from: number): number {
    let i = from;
    while (i < source.length && source[i] !== '\n') {
        i += 1;
    }
    return i < source.length ? i + 1 : i;
}

/** The keyword immediately before `start`, lowercased, across continuations. */
function precedingKeyword(source: string, start: number): string | undefined {
    let i = start - 1;
    for (;;) {
        while (i >= 0 && (source[i] === ' ' || source[i] === '\t')) {
            i -= 1;
        }
        // Step back over a `_` line continuation and the newline before it.
        if (i >= 0 && (source[i] === '\n' || source[i] === '\r')) {
            let j = i;
            while (j >= 0 && (source[j] === '\n' || source[j] === '\r')) {
                j -= 1;
            }
            while (j >= 0 && (source[j] === ' ' || source[j] === '\t')) {
                j -= 1;
            }
            if (j >= 0 && source[j] === '_') {
                i = j - 1;
                continue;
            }
            return undefined;
        }
        break;
    }
    if (i < 0) {
        return undefined;
    }
    let wordEnd = i + 1;
    while (i >= 0 && /[A-Za-z]/.test(source[i])) {
        i -= 1;
    }
    const word = source.slice(i + 1, wordEnd).toLowerCase();
    return word.length > 0 ? word : undefined;
}
