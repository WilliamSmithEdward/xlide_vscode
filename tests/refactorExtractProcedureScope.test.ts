import { describe, expect, it, vi } from 'vitest';
import { extractMethod } from '../src/analyzer/refactor/extractMethod';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
import { VBA_IDENTIFIER_PATTERN } from '../src/vbaSourceScan';

describe.each(['\n', '\r\n', '\r'])('Extract Method procedure scope with %j', eol => {
    it('does not sweep identifiers in unrelated procedures', () => {
        const other = ['Sub Other()', 'Dim foreignAuditVariable As Long', ...Array(1000).fill('foreignAuditVariable = 1'), 'End Sub', ''].join(eol);
        const prefix = 'Option Explicit' + eol + other + ['Sub Main()', 'Dim x As Long', ''].join(eol);
        const body = 'Debug.Print x';
        const source = prefix + body + eol + 'End Sub' + eol + other.replace('Sub Other()', 'Sub Later()');
        const pattern = new RegExp(VBA_IDENTIFIER_PATTERN, 'gu').source, original = RegExp.prototype.exec;
        let foreignSweeps = 0;
        const spy = vi.spyOn(RegExp.prototype, 'exec').mockImplementation(function (this: RegExp, text: string) {
            if (this.source === pattern && this.flags === 'gu' && text.includes('foreignAuditVariable')) { foreignSweeps++; }
            return original.call(this, text);
        });
        let result: ReturnType<typeof extractMethod>;
        try { result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' }); }
        finally { spy.mockRestore(); }
        expect(result!.ok).toBe(true);
        expect(foreignSweeps).toBe(0);
    });

    it.each([
        ['Dim x As Long', 'x = 1', 'Debug.Print x', 'Private Function Work() As Long', 'x = Work()'],
        ['Dim x As Long', 'x = x + 1', 'Debug.Print x', 'Private Sub Work(ByRef x As Long)', 'Work x'],
        ['Dim x As Variant', 'Mutate x', 'Debug.Print x', 'Private Sub Work(ByRef x As Variant)', 'Work x'],
        ['Dim Δ As Long', 'Debug.Print Δ', '', 'Private Sub Work(ByVal Δ As Long)', 'Work Δ'],
    ])('retains caller-local behavior for %s / %s', (decl, body, after, header, invocation) => {
        const convert = (s: string) => s.replace(/\n/g, eol);
        const other = convert('Sub Other()\nDim x As Long, Δ As Long\nx = 4\nΔ = 2\nEnd Sub\n');
        const prefix = 'Option Explicit' + eol + other + convert('Sub Main()\n' + decl + '\n');
        const source = prefix + body + eol + after + eol + 'End Sub' + eol + other.replace('Sub Other()', 'Sub Later()');
        const result = extractMethod({ source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' });
        expect(result.ok).toBe(true); if (!result.ok) throw Error(result.reason);
        const applied = applyVbaTextEdits(source, result.edits);
        expect(applied).toContain(header); expect(applied).toContain(invocation);
        expect(applied).toContain(other); expect(applied).toContain(other.replace('Sub Other()', 'Sub Later()'));
    });
});
