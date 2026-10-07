import { describe, expect, it, vi } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
function fixture(before: string, declarations: string, selected: string, header = 'Sub Main()') {
    const prefix = before + 'Option Explicit\n' + header + '\n' + declarations + '\n';
    return { source: prefix + selected + '\nDebug.Print "after"\nEnd Sub\n', span: { start: prefix.length, end: prefix.length + selected.length }, name: 'Work' };
}
function applied(input: ReturnType<typeof fixture>) {
    const result = extractMethod(input);
    if (!result.ok) { throw new Error(result.reason); }
    return applyVbaTextEdits(input.source, result.edits);
}
describe('Extract Method local ordering', () => {
    it('searches the source at most once per touched local instead of once per comparison', () => {
        const names = Array.from({ length: 1000 }, (_, i) => 'local' + i.toString().padStart(4, '0'));
        const order = names.map((_, i) => names[(i * 37) % names.length]);
        const input = fixture("' " + 'padding '.repeat(1000) + "\n' " + order.join(' ') + '\n',
            names.map(name => 'Dim ' + name + ' As Long').join('\n'), 'Debug.Print ' + names.join(', '));
        const wanted = new Set(names), original = String.prototype.indexOf;
        let searches = 0;
        const spy = vi.spyOn(String.prototype, 'indexOf').mockImplementation(function (this: string, search: string, position?: number) {
            if (String(this) === input.source && wanted.has(search) && position === undefined) { searches++; }
            return original.call(this, search, position);
        });
        let result: ReturnType<typeof extractMethod>;
        try { result = extractMethod(input); } finally { spy.mockRestore(); }
        expect(result!.ok).toBe(true);
        expect(searches).toBeLessThanOrEqual(names.length);
        if (!result!.ok) { throw new Error(result!.reason); }
        expect(applyVbaTextEdits(input.source, result!.edits)).toContain('Private Sub Work(' + order.map(name => 'ByRef ' + name + ' As Long').join(', ') + ')');
    });

    it('preserves raw first occurrence ordering including comments', () => {
        expect(applied(fixture("' Beta Alpha\n", 'Dim Alpha As Long\nDim Beta As Long', 'Debug.Print Alpha, Beta')))
            .toContain('Private Sub Work(ByRef Beta As Long, ByRef Alpha As Long)');
    });

    it('preserves stable ties when the first raw occurrence is a shared substring', () => {
        expect(applied(fixture("' arr\n", 'Dim arr As Long\nDim a As Long', 'Debug.Print a, arr')))
            .toContain('Private Sub Work(ByRef arr As Long, ByRef a As Long)');
    });

    it('keeps procedure parameters before later declarations and handles single locals', () => {
        expect(applied(fixture('', 'Dim later As Long', 'Debug.Print later, first', 'Sub Main(ByRef first As Long)')))
            .toContain('Private Sub Work(ByRef first As Long, ByRef later As Long)');
        expect(applied(fixture('', 'Dim one As Long', 'Debug.Print one')))
            .toContain('Private Sub Work(ByRef one As Long)');
    });
});
