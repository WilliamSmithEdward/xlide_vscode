import { describe, expect, it, vi } from 'vitest';
import * as helpers from '../src/analyzer/lexer/tokenHelpers';
import { standaloneEmptyParenthesizedCallStatement, standaloneMultiArgParenthesizedCallStatement } from '../src/analyzer/call/callContext';

const consumers = [standaloneEmptyParenthesizedCallStatement, standaloneMultiArgParenthesizedCallStatement] as const;
describe('statement-ending parenthesis scan work', () => {
    for (const consumer of consumers) {
        it.each(['nested', 'unclosed'])(consumer.name + ' reads nested tokens linearly: %s', shape => {
            const count = 1000;
            const source = 'F('.repeat(count) + '1' + (shape === 'nested' ? ')'.repeat(count) : '');
            const span = { start: 0, end: source.length };
            const original = helpers.statementTokensCached(source, span);
            let reads = 0;
            const observed = original.map(token => Object.freeze({ ...token, get rawText() { reads++; return token.rawText; } }));
            Object.freeze(observed);
            const spy = vi.spyOn(helpers, 'statementTokensCached').mockReturnValue(observed);
            try {
                expect(consumer(source, span)).toBeUndefined();
                expect(reads).toBeLessThan(original.length * 20);
            } finally { spy.mockRestore(); }
        });
    }

    it('retains the single-query fast path for a shallow call', () => {
        const spy = vi.spyOn(helpers, 'matchParenFrom');
        try {
            expect(standaloneMultiArgParenthesizedCallStatement('F(1, 2)', { start: 0, end: 7 })).toEqual({ name: 'F', isMember: false, qualifier: undefined, argumentCount: 2, span: { start: 0, end: 7 } });
            expect(spy).toHaveBeenCalledTimes(1);
        } finally { spy.mockRestore(); }
    });

    it.each(['F(G(1)).Done()', '.F(G(1)).Done()', 'F()(G(1)).Done()'])('retains complete receiver chain calls: %s', source => {
        const result = standaloneEmptyParenthesizedCallStatement(source, { start: 0, end: source.length });
        expect(result).toEqual({ name: 'Done', isMember: true, startsWithLeadingDot: source.startsWith('.'), calleeEndOffset: source.length - 2, emptyParensSpan: { start: source.length - 2, end: source.length }, span: { start: source.length - 6, end: source.length } });
    });

    it.each(['Call F(G(1))', 'x = F(G(1))', 'F(G(1)) + Done()', 'F((Done()))', 'F[Done()]'])('retains rejected expression and incomplete-chain forms: %s', source => {
        for (const consumer of consumers) expect(consumer(source, { start: 0, end: source.length })).toBeUndefined();
    });
});
