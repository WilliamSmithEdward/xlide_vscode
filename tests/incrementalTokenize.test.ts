import { describe, expect, it } from 'vitest';
import { tokenize, tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { incrementalTokenize } from '../src/analyzer/lexer/incrementalTokenize';
const prefix = "' A stable prefix for the large-module cache\n".repeat(420);
const bodies = [
    'Sub Demo()\nDim value As Long\nvalue = 12\nEnd Sub\n',
    'Sub Demo()\nvalue = "hello"\nvalue = #1/1/2000#\nEnd Sub\n',
    'Sub Demo()\nvalue = Left( _\n"hello", 2)\nEnd Sub\n',
    'Sub Demo()\nOption Compare Text\nOpen "a" For Output As #1\nEnd Sub\n',
    "Sub Demo()\n  ' a comment\n[Unicode α] = 1\nEnd Sub\n",
];

describe('incremental logical-line tokenization', () => {
    it.each(['\n', '\r\n', '\r'])('matches full lexing across edits with %j newlines', newline => {
        let state = 12345;
        const next = () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state; };
        for (const body of bodies) {
            let source = (prefix + body + '  ').replace(/\n/g, newline);
            for (let i = 0; i < 40; i++) {
                const offset = prefix.replace(/\n/g, newline).length + next() % (source.length - prefix.replace(/\n/g, newline).length);
                if (/\r|\n/.test(source[offset])) { continue; }
                const insert = ['', 'x', ' ', '"', "'", '_', ':', '#', 'α', '!', '.'][next() % 11];
                const remove = next() % 2;
                const changed = source.slice(0, offset) + insert + source.slice(offset + remove);
                const previous = tokenize(source);
                const original = JSON.stringify(previous);
                const fast = incrementalTokenize(changed, source, previous, tokenize);
                if (fast) { expect(fast).toEqual(tokenize(changed)); }
                expect(tokenizeCached(changed)).toEqual(tokenize(changed));
                expect(JSON.stringify(previous)).toBe(original);
                source = changed;
            }
        }
    });

    it('re-lexes bounded windows and preserves prefix token identity', () => {
        const source = prefix + bodies[0] + prefix;
        const previous = tokenize(source);
        const changed = source.replace('value = 12', 'value = 123');
        const lengths: number[] = [];
        const result = incrementalTokenize(changed, source, previous, text => { lengths.push(text.length); return tokenize(text); });
        expect(result).toEqual(tokenize(changed));
        expect(result![0]).toBe(previous[0]);
        expect(Math.max(...lengths)).toBeLessThan(100);
    });

    it('extends past absorbed newlines when adding a continuation', () => {
        const source = prefix + 'Sub Demo()\nvalue = 1 \n + 2\nEnd Sub\n';
        const changed = source.replace('1 \n', '1 _\n');
        const result = incrementalTokenize(changed, source, tokenize(source), tokenize);
        expect(result).toBeDefined();
        expect(result).toEqual(tokenize(changed));
    });

    it('falls back for newline changes', () => {
        const source = prefix + bodies[0];
        expect(incrementalTokenize(source.replace('12', '12\nvalue = 2'), source, tokenize(source), tokenize)).toBeUndefined();
    });
});
