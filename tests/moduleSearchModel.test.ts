import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { moduleSearchChangePreview, moduleSearchPreviewer, moduleSearchReplacement, validateModuleSearchMatches } from '../src/moduleSearchModel';

const { searchModule } = require('../assets/webview/moduleSearchWorker.js');
const query = (pattern: string, extra = {}) => ({ pattern, regex: false, wholeWord: false, matchCase: false, ...extra });

describe('module search matching and edits', () => {
    it('accepts dedicated-worker messages and rejects messages bearing a foreign origin', () => {
        const responses: any[] = [];
        const self = { onmessage: undefined as any, postMessage: (message: any) => responses.push(message) };
        runInNewContext(readFileSync('assets/webview/moduleSearchWorker.js', 'utf8'), { self });
        const data = { id: 7, text: 'needle', query: query('needle'), limit: 5 };
        self.onmessage({ origin: 'https://foreign.invalid', data });
        expect(responses).toHaveLength(0);
        self.onmessage({ origin: '', data });
        expect(responses).toMatchObject([{ id: 7, result: { matches: [{ start: 0, end: 6 }] } }]);
    });
    it('previews deletion with unchanged context and multiline replacement with complete matched text', () => {
        const source = 'MsgBox "Passed", vbInformation\r\nnext';
        const match = searchModule(source, query('MsgBox'), 5).matches[0];
        expect(moduleSearchChangePreview(source, match, '')).toEqual({
            prefix: '', suffix: ' "Passed", vbInformation', before: { text: 'MsgBox', truncated: false },
            after: { text: '', truncated: false }, contextTruncated: false,
        });
        const multiline = searchModule('a\r\nb\r\nc', query('a\\r\\nb', { regex: true }), 5).matches[0];
        expect(moduleSearchChangePreview('a\r\nb\r\nc', multiline, '<tag>\n$&')).toMatchObject({
            before: { text: 'a\r\nb' }, after: { text: '<tag>\n$&' },
        });
    });

    it('previews zero-width insertions at the start and explicitly bounds long changes', () => {
        const match = { start: 0, end: 0, captures: [''], named: {} };
        expect(moduleSearchChangePreview('\nnext', match, 'added')).toMatchObject({ prefix: '', before: { text: '' }, after: { text: 'added' } });
        const long = moduleSearchChangePreview('a'.repeat(2000), { ...match, start: 200, end: 1000 }, 'b'.repeat(800));
        expect(long.before).toEqual({ text: 'a'.repeat(500), truncated: true });
        expect(long.after).toEqual({ text: 'b'.repeat(500), truncated: true });
        expect(long.prefix).toHaveLength(100); expect(long.suffix).toHaveLength(120);
        expect(long.contextTruncated).toBe(true);
    });

    it('matches literal punctuation and preserves UTF-16 positions across CRLF and emoji', () => {
        const source = '😀 a.b\r\nA.B';
        const result = searchModule(source, query('a.b'), 50);
        const matches = validateModuleSearchMatches(source, result.matches, 50);
        expect(matches.map(m => [m.start, m.end])).toEqual([[3, 6], [8, 11]]);
        const preview = moduleSearchPreviewer(source);
        expect(matches.map(preview)).toMatchObject([{ line: 1, text: '😀 a.b', highlightStart: 3, highlightEnd: 6 },
            { line: 2, text: 'A.B', highlightStart: 0, highlightEnd: 3 }]);
        expect(searchModule(source, query('a.b', { matchCase: true }), 50).matches).toHaveLength(1);
    });

    it('uses Unicode identifier boundaries for whole-word searches', () => {
        const source = '名foo foo_bar 𐐀foo fooé foo foo';
        expect(searchModule(source, query('foo', { wholeWord: true }), 50).matches).toHaveLength(2);
    });

    it('bounds results and advances zero-length matches by Unicode code point', () => {
        expect(searchModule('a a a', query('a'), 2)).toMatchObject({ limitHit: true, matches: [{ start: 0 }, { start: 2 }] });
        expect(searchModule('😀a', query('(?=.)', { regex: true }), 50).matches.map((m: { start: number }) => m.start)).toEqual([0, 2]);
        expect(() => searchModule('x', query('(', { regex: true }), 50)).toThrow();
    });

    it.each(['$&', '$1-$2', '$12', '$01', '$00', '$99', '$$', '$<first>', '$<missing>', '$<constructor>', '$<__proto__>', '$`', "$'"])(
        'matches native replacement semantics for %s', replacement => {
            const source = '#ab!';
            const match = searchModule(source, query('(?<first>a)(b)?', { regex: true }), 5).matches[0];
            const text = moduleSearchReplacement(source, match, replacement, true);
            expect(source.slice(0, match.start) + text + source.slice(match.end))
                .toBe(source.replace(/(?<first>a)(b)?/u, replacement));
        });

    it('keeps dollar tokens literal for literal replacement and handles unmatched groups', () => {
        const match = searchModule('a', query('(a)(b)?', { regex: true }), 5).matches[0];
        expect(moduleSearchReplacement('a', match, '$2', true)).toBe('');
        expect(moduleSearchReplacement('a', match, '$1', false)).toBe('$1');
    });

    it('bounds replacement expansion before allocating repeated source prefixes', () => {
        const source = 'prefix needle';
        const match = searchModule(source, query('needle'), 5).matches[0];
        expect(() => moduleSearchReplacement(source, match, '$`$`$`', true, 10)).toThrow('size limit');
    });

    it('rejects forged out-of-bounds, mismatching, overlapping or oversized results', () => {
        const valid = searchModule('abc abc', query('abc'), 50).matches;
        expect(() => validateModuleSearchMatches('abc abc', valid, 1)).toThrow();
        for (const altered of [{ ...valid[0], start: -1 }, { ...valid[0], end: 100 },
            { ...valid[0], captures: ['xyz'] }, { ...valid[0], named: null }]) {
            expect(() => validateModuleSearchMatches('abc abc', [altered], 50)).toThrow();
        }
        expect(() => validateModuleSearchMatches('abc abc', [valid[0], valid[0]], 50)).toThrow();
        const empty = searchModule('abc', query('(?=a)', { regex: true }), 5).matches[0];
        expect(() => validateModuleSearchMatches('abc', [empty, empty], 5)).toThrow();
    });
});
