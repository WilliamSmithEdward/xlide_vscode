import { describe, expect, it } from 'vitest';
import { inlineVariable } from '../src/analyzer/refactor/inlineVariable';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { wholeLineSpan } from '../src/vbaSourceScan';

const cases = [
    ['following neighbors', 'Dim limit As Long: Debug.Print "decl"', 'limit = 3: Debug.Print "assign"', [' Debug.Print "decl"', ' Debug.Print "assign"']],
    ['preceding neighbors', 'Debug.Print "decl": Dim limit As Long', 'Debug.Print "assign": limit = 3', ['Debug.Print "decl"', 'Debug.Print "assign"']],
    ['labels', 'D: Dim limit As Long', 'A: limit = 3', ['D: ', 'A: ']],
    ['comments', "Dim limit As Long ' declaration", "limit = 3 ' assignment", [" ' declaration", " ' assignment"]],
    ['quoted colon', 'Dim limit As Long', 'limit = 3: Debug.Print "a:b"', [' Debug.Print "a:b"']],
    ['one physical line', 'Dim limit As Long: limit = 3: Debug.Print limit', '', ['  Debug.Print VALUE']],
    ['adjacent deletions', 'Debug.Print "keep": Dim limit As Long: limit = 3', '', ['Debug.Print "keep": ']],
    ['label and use', 'Dim limit As Long', 'L: limit = 3: Debug.Print limit', ['L:  Debug.Print VALUE']],
] as const;

describe('refactoring statement removals', () => {
    for (const eol of ['\n', '\r\n', '\r']) for (const [name, declaration, assignment, surviving] of cases) {
        const source = ['Sub P()', declaration, ...(assignment ? [assignment] : []), 'Debug.Print limit', 'End Sub', ''].join(eol);
        for (const mode of ['inline', 'parameter'] as const) {
            it(mode + ' preserves ' + name + ' with ' + JSON.stringify(eol), () => {
                const offset = source.indexOf('limit');
                const result = mode === 'inline' ? inlineVariable({ source, offset }) : introduceParameter({ source, offset, moduleName: 'M' });
                expect(result.ok).toBe(true);
                if (!result.ok) throw new Error(result.reason);
                const value = mode === 'inline' ? '3' : 'limit';
                const expected = [mode === 'inline' ? 'Sub P()' : 'Sub P(ByVal limit As Long)', ...surviving.map(line => line.replace('VALUE', value)), 'Debug.Print ' + value, 'End Sub', ''].join(eol);
                expect(applyVbaTextEdits(source, result.edits)).toBe(expected);
                const ordered = [...result.edits].sort((a, b) => a.span.start - b.span.start);
                for (let i = 1; i < ordered.length; i++) expect(ordered[i].span.start).toBeGreaterThanOrEqual(ordered[i - 1].span.end);
            });
        }
    }

    it.each(['\n', '\r\n'])('preserves every legacy LF/CRLF span boundary with %j', eol => {
        const source = ['first', 'middle', 'last'].join(eol);
        for (let start = 0; start <= source.length; start++) for (let end = start; end <= source.length; end++) {
            const after = source.indexOf('\n', end);
            expect(wholeLineSpan(source, { start, end })).toEqual({ start: source.lastIndexOf('\n', start - 1) + 1, end: after < 0 ? source.length : after + 1 });
        }
    });
    it('widens only the selected CR-only physical line', () => {
        const source = 'first\rmiddle\rlast';
        expect(wholeLineSpan(source, { start: 7, end: 10 })).toEqual({ start: 6, end: 13 });
    });
});
