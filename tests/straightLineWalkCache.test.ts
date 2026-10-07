import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { straightLineAssignments, straightLineExit, straightLineUnreachable } from '../src/analyzer/diagnostics/straightLineValues';

const tokens = (text: string) => tokenize(text).filter(token => token.kind !== 'eof');

describe('straight-line walk cache', () => {
    it('reuses equal starts in different insertion orders and distinguishes different values', () => {
        const source = 'Sub P()\nIf n = 0 Then Exit Sub\nDebug.Print n\nEnd Sub';
        const mod = parseModule(source);
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const first = new Map([['n', tokens('0')], ['other', tokens('1')]]);
        const equal = new Map([['other', tokens('1')], ['n', tokens('0')]]);
        const different = new Map([['n', tokens('1')], ['other', tokens('1')]]);
        const dead = straightLineUnreachable(source, proc.body, undefined, first);
        expect(dead.has(proc.body[1])).toBe(true);
        expect(straightLineUnreachable(source, proc.body, undefined, equal)).toBe(dead);
        expect(straightLineUnreachable(source, proc.body, undefined, different).has(proc.body[1])).toBe(false);
        expect(straightLineUnreachable(source, proc.body, undefined, first)).toBe(dead);
    });

    it('invalidates an identical start when conditional activity changes', () => {
        const source = 'Sub P()\n#If FLAG Then\nn = 1\n#Else\nn = 2\n#End If\nEnd Sub';
        const mod = parseModule(source);
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const initial = new Map([['n', tokens('0')]]);
        for (const flag of [true, false, true]) {
            const activity = createConditionalActivityTracker(mod, { compilerConstants: { FLAG: flag } });
            const result = straightLineExit(source, proc.body, activity, initial);
            expect(result.exit?.get('n')?.map(token => token.rawText).join('')).toBe(flag ? '1' : '2');
        }
    });

    it('does not serialize constant values when a body only uses one start', () => {
        let reads = 0;
        const literal = tokens('1')[0];
        const counted = { ...literal };
        Object.defineProperty(counted, 'rawText', { get: () => { reads++; return literal.rawText; } });
        const initial = new Map(Array.from({ length: 500 }, (_, i) => [`k${i}`, [counted]] as const));
        const body: [] = [];
        const result = straightLineAssignments('', body, undefined, initial);
        expect(straightLineAssignments('', body, undefined, initial)).toBe(result);
        expect(reads).toBe(0);
    });

    it('invalidates an identical start when the source changes on a reused body', () => {
        const source = 'Sub P()\nn = 1\nEnd Sub';
        const mod = parseModule(source);
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const initial = new Map([['n', tokens('0')]]);
        for (const value of [1, 2, 1]) {
            const text = source.replace('n = 1', `n = ${value}`);
            expect(straightLineExit(text, proc.body, undefined, initial).exit?.get('n')?.[0].rawText).toBe(String(value));
        }
    });

    it('compares shared constant tokens without formatting separately allocated starts', () => {
        let reads = 0;
        const literal = tokens('1')[0];
        const counted = { ...literal };
        Object.defineProperty(counted, 'rawText', { get: () => { reads++; return literal.rawText; } });
        const value = [counted];
        const entries = Array.from({ length: 500 }, (_, i) => [`k${i}`, value] as const);
        const body: [] = [];
        const first = straightLineAssignments('', body, undefined, new Map(entries));
        for (let i = 0; i < 5; i++) {
            expect(straightLineAssignments('', body, undefined, new Map(entries))).toBe(first);
        }
        expect(reads).toBe(0);
    });

    it('adopts an equal start so repeated queries do not compare its entries again', () => {
        let lookups = 0;
        class CountedStart extends Map<string, ReturnType<typeof tokens>> {
            override get(name: string) { lookups++; return super.get(name); }
        }
        const entries = Array.from({ length: 500 }, (_, i) => [`k${i}`, tokens('1')] as const);
        const initial = new CountedStart(entries);
        const equal = new Map([...entries].reverse());
        const body: [] = [];
        const result = straightLineAssignments('', body, undefined, initial);
        expect(straightLineAssignments('', body, undefined, equal)).toBe(result);
        expect(lookups).toBe(500);
        for (let i = 0; i < 10; i++) expect(straightLineAssignments('', body, undefined, equal)).toBe(result);
        expect(lookups).toBe(500);
    });
    it('discards stale snapshots before formatting starts from another source', () => {
        let reads = 0;
        const literal = tokens('1')[0];
        const counted = { ...literal };
        Object.defineProperty(counted, 'rawText', { get: () => { reads++; return literal.rawText; } });
        const start = new Map(Array.from({ length: 500 }, (_, i) => [`k${i}`, [counted]] as const));
        const body: [] = [];
        const first = straightLineAssignments('first source', body, undefined, start);
        expect(straightLineAssignments('second source', body, undefined, start)).not.toBe(first);
        expect(reads).toBe(0);
    });

});
