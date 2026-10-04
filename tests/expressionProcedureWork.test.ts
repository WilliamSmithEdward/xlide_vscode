import { describe, expect, it, vi } from 'vitest';
import * as parser from '../src/analyzer/parser/parseModule';
import { resolveExpressionType } from '../src/analyzer/expression/resolveExpressionType';

describe('expression enclosing procedure work', () => {
    it.each([100, 1000])('does not rescan %i preceding members on each warm query', count => {
        const source = Array.from({ length: count }, (_, i) => 'Function F' + i + '() As Long\nF' + i + ' = 1\nEnd Function\n').join('');
        const start = source.lastIndexOf('1'), span = { start, end: start + 1 };
        const module = parser.parseModule(source);
        let reads = 0;
        for (const member of module.members) {
            const kind = member.kind;
            Object.defineProperty(member, 'kind', { get() { reads++; return kind; } });
            Object.freeze(member);
        }
        Object.freeze(module.members);
        const spy = vi.spyOn(parser, 'parseModule').mockReturnValue(module);
        try {
            const ctx = { moduleName: 'Work' + count };
            expect(resolveExpressionType(source, span, ctx)).toEqual({ type: 'Long', isObject: false, complete: true });
            reads = 0;
            for (let query = 0; query < 100; query++) expect(resolveExpressionType(source, span, ctx)).toEqual({ type: 'Long', isObject: false, complete: true });
            expect(reads).toBeLessThan(count * 3);
        } finally { spy.mockRestore(); }
    });

    it('keeps the original first-containing procedure when malformed intervals overlap', () => {
        const source = 'Sub First()\nDim x As Long\nx = 1\nEnd Sub\nSub Second()\nDim x As String\nx = "s"\nEnd Sub\n';
        const module = parser.parseModule(source);
        module.members[0].span.end = source.length;
        for (const member of module.members) { Object.freeze(member.span); Object.freeze(member); }
        Object.freeze(module.members);
        const spy = vi.spyOn(parser, 'parseModule').mockReturnValue(module);
        try {
            const start = source.lastIndexOf('x =');
            for (let i = 0; i < 4; i++) expect(resolveExpressionType(source, { start, end: start + 1 }, { moduleName: 'Overlapping' })).toEqual({ type: 'Long', isObject: false, complete: true });
        } finally { spy.mockRestore(); }
    });

    it('keeps distinct source and module contexts from sharing a procedure index', () => {
        const source = 'Sub P()\nDim x As String\nx = "s"\nEnd Sub\n';
        const start = source.lastIndexOf('x =');
        for (const moduleName of ['One', 'Two', 'One']) for (let query = 0; query < 3; query++) expect(resolveExpressionType(source, { start, end: start + 1 }, { moduleName })).toEqual({ type: 'String', isObject: false, complete: true });
        const changed = source.replace('Dim x As String', 'Dim x As Collection');
        const changedStart = changed.lastIndexOf('x =');
        for (let query = 0; query < 3; query++) expect(resolveExpressionType(changed, { start: changedStart, end: changedStart + 1 }, { moduleName: 'One' })).toEqual({ type: 'Collection', isObject: true, complete: true });
    });
});
