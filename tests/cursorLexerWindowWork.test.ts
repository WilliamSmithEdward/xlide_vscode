import { afterEach, describe, expect, it, vi } from 'vitest';
import { completionCursorContext, completionLineCursorContext } from '../src/analyzer/completion/cursorContext';
import { resolveTypeCompletions } from '../src/analyzer/completion/typeCompletion';
import * as lexer from '../src/analyzer/lexer/tokenize';

afterEach(() => vi.restoreAllMocks());

describe('fresh cursor lexer windows', () => {
    const fixtures = [
        'x = value + value',
        'Dim item As _\n New _\n Shapes.Person',
        "Rem note _\n next comment _\nlast comment",
        'ThisWorkbook.Sheets(1). _\n Name = "a ""quoted"" string"',
        "x = 1: y = [odd name] ' comment",
        'Open "p" For Output As #1',
        'x = "broken _\nNextLine',
        'x = #1/1/2026#',
        'x = 1 _\u0019\n+2',
        "' note _\u3000\nmore _\nend",
        'abc_\nDef',
        'x = 1   \n  y = Other',
        ' \n \n Dim Value As Long\n  ',
        '[broken\nAs New Person',
        'If ok Then Rem note\nobj.Rem _\n  .Name',
        '#If ready Then\nDim Value As Long\n#End If',
        '\n', '',
    ];

    it.each(['\n', '\r\n', '\r'])('preserves complete token metadata and cursor decisions with %j breaks', ending => {
        for (const body of fixtures) {
            const source = ('Option Explicit\nSub Padding()\nDebug.Print 1\nEnd Sub\n' + body).replaceAll('\n', ending);
            for (let offset = 0; offset <= source.length; offset++) {
                const full = completionCursorContext(source, offset);
                const line = completionLineCursorContext(source, offset);
                const expected = line.tokens.length ? full.tokens.slice(-line.tokens.length) : [];
                expect(line.tokens, 'token metadata at ' + offset).toEqual(expected);
                expect(line.partialToken).toEqual(full.partialToken);
                expect(line.before).toEqual(full.before);
                expect(line.partial).toBe(full.partial);
                expect(line.statementStart).toBe(full.statementStart);
                expect(line.inComment).toBe(full.inComment);
                expect(line.inString).toBe(full.inString);
            }
        }
    });

    it.each(['head', 'tail'])('bounds actual cold lexer input on fresh revisions near the %s', position => {
        const padding = 'Sub ColdRevision()\r\nDebug.Print 1\r\nEnd Sub\r\n'.repeat(1400);
        const probes = [
            { source: 'value = other + value', type: undefined },
            { source: 'value', type: undefined }, // Fewer than five grammar tokens: inspect preceding lines.
            { source: 'ThisWorkbook.Sheets(1).Na', type: undefined },
            { source: 'Dim item As Lon', type: 'Long' },
            { source: 'Dim item As New Shapes.Pe', type: 'Person' },
            { source: 'Dim item As _\r\n New _\n Shapes.Pe', type: 'Person' },
            { source: "Dim item As\n' note _\u3000\r\n New _\n Shapes.Pe", type: 'Person' },
        ];
        for (const [revision, probe] of probes.entries()) {
            const lead = "' cold revision " + revision + '\r\n';
            const source = position === 'head' ? lead + probe.source + '\r\n' + padding
                : padding + lead + probe.source;
            const offset = position === 'head' ? lead.length + probe.source.length : source.length;
            const original = lexer.tokenize;
            const inputs: number[] = [];
            const spy = vi.spyOn(lexer, 'tokenize').mockImplementation(text => {
                inputs.push(text.length);
                return original(text);
            });
            lexer.startTokenizeMissLogForTests();
            let misses: number[];
            try {
                const cursor = completionLineCursorContext(source, offset);
                expect(cursor.offset).toBe(offset);
                const types = resolveTypeCompletions(source, offset, {
                    projectTypes: [{ name: 'Person', kind: 'class', moduleName: 'Shapes' }],
                });
                if (probe.type) expect(types.map(item => item.name)).toContain(probe.type);
                else expect(types).toEqual([]);
            } finally {
                spy.mockRestore();
                misses = lexer.stopTokenizeMissLogForTests();
            }
            expect(misses).toEqual([]);
            expect(inputs.length).toBeGreaterThan(0);
            expect(Math.max(...inputs)).toBeLessThan(200);
            expect(inputs.reduce((a, b) => a + b, 0)).toBeLessThan(500);
        }
    });

    it('preserves absolute line numbers when a consumer explicitly requests them', () => {
        const source = 'Sub MetadataProbe()\r\nDebug.Print 1\nEnd Sub\r'.repeat(1200) + 'value = other + value';
        const context = completionLineCursorContext(source, source.length);
        expect(context.partialToken?.line).toBe(3600);
        expect(context.partialToken?.character).toBe(16);
        expect(context.tokens).toEqual(completionCursorContext(source, source.length).tokens.slice(-context.tokens.length));
        expect(JSON.parse(JSON.stringify(context.partialToken))).toEqual({
            kind: 'identifier', rawText: 'value', start: source.length - 5, end: source.length,
            line: 3600, character: 16,
            leadingTrivia: [{ kind: 'whitespace', text: ' ', start: source.length - 6, end: source.length - 5 }],
        });
    });
});

it('does not independently re-lex thousands of sparse type lookback rows', () => {
    const source = 'Dim item As\n' + "' sparse comment\r\n \n".repeat(2200) + 'New Shapes.Pe';
    const tokenize = lexer.tokenize;
    const inputs: number[] = [];
    const spy = vi.spyOn(lexer, 'tokenize').mockImplementation(text => {
        inputs.push(text.length);
        return tokenize(text);
    });
    const shared = vi.spyOn(lexer, 'tokenizeCached');
    try {
        expect(resolveTypeCompletions(source, source.length, {
            projectTypes: [{ name: 'Person', kind: 'class', moduleName: 'Shapes' }],
        }).map(item => item.name)).toEqual(['Person']);
    } finally {
        spy.mockRestore();
    }
    expect(inputs.length).toBeLessThan(100);
    expect(shared).toHaveBeenCalledWith(source);
});
