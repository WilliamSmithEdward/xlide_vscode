import type { Trivia, VbaToken } from './tokenKinds';

function shifted(token: VbaToken, delta: number): VbaToken {
    if (!delta) { return token; }
    const trivia = (items: Trivia[] | undefined) => items?.map(item => ({ ...item, start: item.start + delta, end: item.end + delta }));
    return { ...token, start: token.start + delta, end: token.end + delta,
        ...(token.leadingTrivia ? { leadingTrivia: trivia(token.leadingTrivia) } : {}),
        ...(token.trailingTrivia ? { trailingTrivia: trivia(token.trailingTrivia) } : {}),
    };
}

/** Re-lex a small same-physical-line edit between unchanged logical boundaries. */
export function incrementalTokenize(source: string, previousSource: string, previous: readonly VbaToken[], lex: (source: string) => VbaToken[]): VbaToken[] | undefined {
    if (source.length < 16_384 || Math.abs(source.length - previousSource.length) > 256 ||
        source.slice(0, 128) !== previousSource.slice(0, 128) || !previous.length) { return undefined; }
    let start = 128;
    const commonEnd = Math.min(source.length, previousSource.length);
    while (start < commonEnd && source.charCodeAt(start) === previousSource.charCodeAt(start)) { start++; }
    let oldEnd = previousSource.length, newEnd = source.length;
    while (oldEnd > start && newEnd > start && previousSource.charCodeAt(oldEnd - 1) === source.charCodeAt(newEnd - 1)) { oldEnd--; newEnd--; }
    if (oldEnd - start > 256 || newEnd - start > 256 || /[\r\n]/.test(previousSource.slice(start, oldEnd) + source.slice(start, newEnd))) { return undefined; }
    let lo = 0, hi = previous.length;
    while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (previous[mid].end < start) { lo = mid + 1; } else { hi = mid; }
    }
    let head = Math.min(lo, previous.length - 1);
    while (head > 0 && previous[head].kind !== 'newline') { head--; }
    const first = previous[head];
    const base = first.leadingTrivia?.[0]?.start ?? first.start;
    // A window can start in a continued statement's trivia. Its absolute line
    // origin then differs from the token's line; leave that case to full lexing.
    if (first.leadingTrivia?.some(item => item.kind === 'lineContinuation')) { return undefined; }
    const baseCharacter = first.character - (first.start - base);
    const delta = source.length - previousSource.length;
    let tail = Math.max(head, lo);
    let boundaries = 0;
    while (tail < previous.length) {
        const boundary = previous[tail++];
        if (boundary.kind !== 'newline' || boundary.end <= oldEnd) { continue; }
        if (tail === previous.length) { break; } // Preserve EOF trailing trivia.
        if (++boundaries > 4 || boundary.end - base > 16_384) { return undefined; }
        const window = lex(source.slice(base, boundary.end + delta));
        const last = window.at(-1);
        // Appending/deleting a continuation can absorb the old newline. Extend
        // to another logical boundary until lexical state converges.
        if (last?.kind !== 'newline' || last.end !== boundary.end + delta - base) { continue; }
        const middle = window.map(token => {
            const rebased = shifted(token, base);
            return { ...rebased, line: token.line + first.line,
                character: token.character + (token.line === 0 ? baseCharacter : 0) };
        });
        return [...previous.slice(0, head), ...middle, ...previous.slice(tail).map(token => shifted(token, delta))];
    }
    // No following logical boundary (editing the final line): lex to EOF.
    if (previousSource.length - base > 16_384) { return undefined; }
    const middle = lex(source.slice(base)).map(token => ({ ...shifted(token, base), line: token.line + first.line,
        character: token.character + (token.line === 0 ? baseCharacter : 0) }));
    return [...previous.slice(0, head), ...middle];
}
