import { describe, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { statementTokensCached } from '../src/analyzer/lexer/tokenHelpers';

let eviction = 0;
function evict() {
    for (let i = 0; i < 3; i++) statementTokensCached("' eviction " + eviction++, { start: 0, end: 1 });
}
const prefix = Array.from({ length: 200 }, (_, i) => 'Dim item' + i + ' As Long\r\n').join('');
const tail = ' result = obj.Name: nextValue = 1\r\n text = "a" _\r\n & "b"\r\n';
function spans(source: string) {
    const result: { start: number; end: number }[] = [];
    let start = 0;
    for (const match of source.matchAll(/\r\n|\r|\n/g)) {
        result.push({ start, end: match.index });
        start = match.index + match[0].length;
    }
    result.push({ start, end: source.length });
    return result;
}
function read(source: string) { return spans(source).map(span => statementTokensCached(source, span)); }

describe('statement token prefix reuse', () => {
    it('reuses complete prefix arrays without module-token queries and keeps snapshots immutable', () => {
        evict();
        const original = prefix + tail;
        const oldSpans = spans(original);
        const before = read(original);
        const snapshot = structuredClone(before);
        const changed = original.replace('obj.Name', 'obj.Value');
        const spy = vi.spyOn(lexer, 'tokenizeCached');
        try {
            for (let i = 0; i < 200; i++) expect(statementTokensCached(changed, oldSpans[i])).toBe(before[i]);
            expect(spy).not.toHaveBeenCalled();
            expect(statementTokensCached(changed, spans(changed)[200])).not.toBe(before[200]);
            expect(spy).toHaveBeenCalled();
        } finally { spy.mockRestore(); }
        expect(before).toEqual(snapshot);
    });

    it('matches cold derived tokens for every tail insertion and deletion, including trivia and positions', () => {
        const original = prefix + tail;
        for (let offset = prefix.length; offset <= original.length; offset++) {
            for (const inserted of ['x', "'", '"', '_', ':', '\r', '\n', 'ก้', '']) {
                if (!inserted && offset === original.length) continue;
                evict(); read(original);
                const changed = original.slice(0, offset) + inserted + original.slice(offset + (inserted ? 0 : 1));
                const incremental = read(changed);
                evict();
                expect(incremental, 'edit ' + offset + ' ' + JSON.stringify(inserted)).toEqual(read(changed));
            }
        }
    }, 60000);

    it('does not share changed-line spans or same-line suffixes', () => {
        evict();
        const original = prefix + 'aaa = 1: bbb = 2\r\n';
        const start = prefix.length;
        const left = { start, end: start + 7 }, right = { start: start + 9, end: start + 16 };
        const a = statementTokensCached(original, left), b = statementTokensCached(original, right);
        const changed = original.replace('bbb = 2', 'bbb = 3');
        expect(statementTokensCached(changed, left)).not.toBe(a);
        expect(statementTokensCached(changed, right)).not.toBe(b);
        expect(statementTokensCached(changed, right).at(-1)?.rawText).toBe('3');
    });

    it('falls back for early changes, large changes and unrelated snapshots', () => {
        for (const changed of ['X' + prefix.slice(1), prefix + 'x'.repeat(129), 'y'.repeat(prefix.length)]) {
            evict();
            const span = { start: 0, end: 17 };
            const old = statementTokensCached(prefix, span);
            expect(statementTokensCached(changed, span)).not.toBe(old);
            const result = statementTokensCached(changed, span);
            evict(); expect(result).toEqual(statementTokensCached(changed, span));
        }
    });
});
