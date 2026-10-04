import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { procedureAtOffset, type ModuleNode, type ProcedureNode } from '../src/analyzer/parser/nodes';
import { resolveHostMemberKindAt } from '../src/analyzer/completion/memberAccess';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

function linear(module: ModuleNode, offset: number) {
    return module.members.find((member): member is ProcedureNode =>
        member.kind === 'Procedure' && offset >= member.span.start && offset <= member.span.end);
}

describe('procedure span lookup', () => {
    it('matches first-match lookup throughout valid and unfinished source', () => {
        for (const source of [
            "' header\r\n" + Array.from({ length: 40 }, (_, i) => 'Sub P' + i + '()\r\nWith ThisWorkbook\r\n .Save\r\nEnd With\r\nEnd Sub\r\n').join(''),
            'Sub OpenOne()\n Dim x As Long\nSub NextOne()\nEnd Sub\n',
            'Function F() As Long\r F = 1\rEnd Function\r\rPublic x As Long\r',
        ]) {
            const module = parseModule(source);
            for (let offset = -1; offset <= source.length + 1; offset++) {
                expect(procedureAtOffset(module, offset)).toBe(linear(module, offset));
            }
        }
    });

    it('reads module members once for thousands of later position lookups', () => {
        const module = structuredClone(parseModule(Array.from({ length: 1500 }, (_, i) =>
            'Sub P' + i + '()\r\nDebug.Print 1\r\nEnd Sub\r\n').join('')));
        let reads = 0;
        module.members = new Proxy(module.members, { get(target, key, receiver) {
            if (typeof key === 'string' && /^\d+$/.test(key)) { reads++; }
            return Reflect.get(target, key, receiver);
        } });
        procedureAtOffset(module, module.span.end - 5);
        expect(reads).toBeGreaterThan(1000);
        reads = 0;
        for (let offset = 0; offset < module.span.end; offset += 11) {
            procedureAtOffset(module, offset);
        }
        expect(reads).toBe(0);
    });

    it('preserves first-match behavior for overlapping, adjacent, reversed and invalid spans', () => {
        for (const spans of [[[10, 100], [20, 30]], [[10, 20], [20, 30]], [[20, 30], [10, 15]], [[20, 10], [5, 30]], [[NaN, 30], [5, 30]], [[10, Infinity], [20, 30]]]) {
            const module = structuredClone(parseModule('Sub A()\nEnd Sub\nSub B()\nEnd Sub'));
            module.members.forEach((member, i) => { member.span = { start: spans[i][0], end: spans[i][1] }; });
            for (let offset = -1; offset < 110; offset++) {
                expect(procedureAtOffset(module, offset)).toBe(linear(module, offset));
            }
            expect(procedureAtOffset(module, NaN)).toBeUndefined();
        }
    });

    it('keeps old spans immutable and uses a new index after source edits', () => {
        const before = parseModule('Sub Before()\nEnd Sub');
        const snapshot = structuredClone(before);
        expect(procedureAtOffset(before, 5)?.name).toBe('Before');
        const after = parseModule("' inserted line\nSub After()\nEnd Sub");
        expect(procedureAtOffset(after, 5)).toBeUndefined();
        expect(procedureAtOffset(after, after.span.end - 2)?.name).toBe('After');
        expect(before).toEqual(snapshot);
    });
});

describe('receiver declaration indexes', () => {
    it('does not walk the same local declaration body on each receiver lookup', () => {
        const source = 'Sub Probe()\r\n' + Array.from({ length: 100 }, (_, i) => ' Dim padding' + i + ' As Long\r\n').join('') +
            ' Dim sheet As Worksheet\r\n sheet.Calculate\r\nEnd Sub';
        const module = structuredClone(parseModule(source));
        const procedure = module.members.find((member): member is ProcedureNode => member.kind === 'Procedure')!;
        let iterations = 0;
        procedure.body = new Proxy(procedure.body, { get(target, key, receiver) {
            if (key === Symbol.iterator) { iterations++; }
            return Reflect.get(target, key, receiver);
        } });
        const context = { parsedModule: module, sourceTokens: tokenizeCached(source) };
        const offset = source.indexOf('sheet.Calculate') + 'sheet.Calculate'.length;
        expect(resolveHostMemberKindAt(source, offset, 'Calculate', context)).toBe('method');
        const first = iterations;
        expect(first).toBeGreaterThan(0);
        for (let i = 0; i < 50; i++) {
            expect(resolveHostMemberKindAt(source, offset, 'Calculate', context)).toBe('method');
        }
        expect(iterations).toBe(first);
    });

    it('preserves parameter/local/module precedence, untyped shadows and first duplicate declarations', () => {
        const cases = [
            ['Public sheet As Workbook\nSub Probe()\n Dim sheet As Worksheet\n sheet.Calculate\nEnd Sub', 'sheet.Calculate', 'Calculate', 'method'],
            ['Public sheet As Workbook\nSub Probe()\n sheet.Close\nEnd Sub', 'sheet.Close', 'Close', 'method'],
            ['Sub Probe(Application As Range)\n Dim Application As Workbook\n Application.Calculate\nEnd Sub', 'Application.Calculate', 'Calculate', 'method'],
            ['Sub Probe()\n Dim sheet As Worksheet\n Dim sheet As Long\n sheet.Calculate\nEnd Sub', 'sheet.Calculate', 'Calculate', 'method'],
            ['Sub Probe()\n Dim Application\n Application.Quit\nEnd Sub', 'Application.Quit', 'Quit', undefined],
            ['Sub Probe()\n For i = 1 To 2\n Dim sheet As Worksheet\n Next\n sheet.Calculate\nEnd Sub', 'sheet.Calculate', 'Calculate', 'method'],
        ] as const;
        for (const [source, text, member, expected] of cases) {
            const context = { parsedModule: parseModule(source), sourceTokens: tokenizeCached(source), allowSetAssignmentRefinement: false };
            for (let i = 0; i < 3; i++) {
                expect(resolveHostMemberKindAt(source, source.indexOf(text) + text.length, member, context), source).toBe(expected);
            }
        }
    });

    it('refreshes declaration types after edits without modifying earlier ASTs', () => {
        const source = 'Sub Probe()\n Dim sheet As Worksheet\n sheet.Calculate\nEnd Sub';
        const oldModule = parseModule(source);
        const snapshot = structuredClone(oldModule);
        const run = (text: string) => resolveHostMemberKindAt(text, text.indexOf('sheet.Calculate') + 'sheet.Calculate'.length,
            'Calculate', { parsedModule: parseModule(text), sourceTokens: tokenizeCached(text) });
        expect(run(source)).toBe('method');
        expect(run(source.replace('As Worksheet', 'As Long'))).toBeUndefined();
        expect(run(source)).toBe('method');
        expect(oldModule).toEqual(snapshot);
    });
});
