import { afterEach, describe, expect, it, vi } from 'vitest';
import { completionCursorContext, completionLineCursorContext, completionTypeTokens } from '../src/analyzer/completion/cursorContext';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { resolveTypeCompletions } from '../src/analyzer/completion/typeCompletion';
import { resolveMemberCompletions } from '../src/analyzer/completion/memberAccess';
import { resolveKeywordCompletions } from '../src/analyzer/completion/keywordCompletion';

afterEach(() => vi.restoreAllMocks());

const snapshot = (tokens: readonly { kind: string; rawText: string; start: number; end: number }[]) => tokens.map(({ kind, rawText, start, end }) => ({ kind, rawText, start, end }));

describe('bounded completion prefix windows', () => {
    const fixture = [
        'Attribute VB_Name = "Módulo"', 'Option Explicit',
        'Sub Пример()', 'Dim item As _', ' New _', ' Shapes.Person',
        'ThisWorkbook.Sheets(1). _', ' Name = "a ""quoted"" string"',
        "Rem continued comment _", ' next comment line',
        "x = 1: y = [odd name] ' comment", 'If x <= 42 Then',
        'End If', 'End Sub', '',
    ].join('\r\n');

    it('matches fresh truncated prefix tokens at every offset, including continuation gaps', () => {
        for (let offset = 0; offset <= fixture.length; offset++) {
            const direct = lexer.tokenize(fixture.slice(0, offset));
            expect(snapshot(completionTypeTokens(fixture, offset)), 'type offset ' + offset)
                .toEqual(snapshot(direct.filter(token => token.kind !== 'comment' && token.kind !== 'newline').slice(-5)));
            const full = completionCursorContext(fixture, offset);
            const line = completionLineCursorContext(fixture, offset);
            expect(snapshot(line.tokens), 'line offset ' + offset).toEqual(snapshot(direct.slice(direct.length - line.tokens.length)));
            expect(line.partialToken).toEqual(full.partialToken);
            expect(line.partial).toEqual(full.partial);
            expect(line.statementStart).toEqual(full.statementStart);
            expect(line.inComment).toBe(full.inComment);
            expect(line.inString).toBe(full.inString);
        }
    });

    it('does not project a module-sized prefix for line or type queries', () => {
        const source = 'Sub Padding()\nDebug.Print 1\nEnd Sub\n'.repeat(1200) + 'ThisWorkbook.Sheets(1).Na';
        const all = lexer.tokenizeCached(source);
        let reads = 0;
        const counted = new Proxy(all, {
            get(target, property, receiver) {
                if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                return Reflect.get(target, property, receiver);
            },
        });
        vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(counted);
        for (let offset = source.length - 5; offset <= source.length; offset++) {
            reads = 0;
            completionLineCursorContext(source, offset);
            completionTypeTokens(source, offset);
            expect(reads).toBeLessThan(100);
        }
    });

    it('preserves continued member chains, colon boundaries and qualified declaration types', () => {
        const member = 'Sub Probe()\nThisWorkbook.Sheets(1). _\r\n Na';
        expect(resolveMemberCompletions(member, member.length).map(item => item.name)).toContain('Name');
        const declaration = 'Dim item As _\r\n New _\r\n Shapes.Pe';
        expect(resolveTypeCompletions(declaration, declaration.length, {
            projectTypes: [{ name: 'Person', kind: 'class', moduleName: 'Shapes' }],
        }).map(item => item.name)).toEqual(['Person']);
        const keyword = 'Sub Probe()\nx = 1: On Error ';
        expect(resolveKeywordCompletions(keyword, keyword.length).items.map(item => item.label)).toContain('Resume Next');
    });
});

it.skipIf(process.env.XLIDE_PREFIX_WINDOW_BENCH !== '1')('measures completion queries at varying positions near the end of a large module', () => {
    const source = Array.from({ length: 1200 }, (_, i) => 'Sub Pad' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('') +
        'Sub Probe()\n' + 'ThisWorkbook.Sheets(1).Name\n'.repeat(6) + 'End Sub\n';
    const offsets = [...source.matchAll(/Sheets\(1\)\.Na/g)].map(match => match.index! + match[0].length);
    const medianMs: Record<string, number> = {};
    for (const [name, run] of [
        ['type', (offset: number) => resolveTypeCompletions(source, offset)],
        ['member', (offset: number) => resolveMemberCompletions(source, offset)],
        ['keyword', (offset: number) => resolveKeywordCompletions(source, offset)],
        ['memberHover', (offset: number) => resolveHover(source, offset)],
    ] as const) {
        for (const offset of offsets) { run(offset); }
        const times = Array.from({ length: 21 }, () => {
            const before = performance.now();
            for (let i = 0; i < 60; i++) { run(offsets[i % offsets.length]); }
            return (performance.now() - before) / 60;
        }).sort((a,b) => a-b);
        medianMs[name] = times[10];
    }
    process.stdout.write('Prefix window benchmark: ' + JSON.stringify({ bytes: source.length, positions: offsets.length, medianMs }) + '\n');
});
